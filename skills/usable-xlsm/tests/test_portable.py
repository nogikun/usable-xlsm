from __future__ import annotations

from io import BytesIO, StringIO
from contextlib import redirect_stdout
import json
import random
import shutil
import struct
import tempfile
from types import ModuleType, SimpleNamespace
from pathlib import Path
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from olefile import OleFileIO
from oletools.olevba import decompress_stream
from typer.testing import CliRunner

from usable_xlsm.cli import app
from usable_xlsm.core import VbaModuleMismatchError, extract_vba, validate_vba_modules
from usable_xlsm.creation import create_workbook, extract_project
from usable_xlsm.environment import capabilities
from usable_xlsm.planning import plan_changes, portability_issues
from usable_xlsm.runner import excel_pids, run_job
from usable_xlsm.testing import assert_module_source, discover, generate_runner
from usable_xlsm.vba_project import _compound_file, _project_property, compress_vba, project_bytes


class PortableTests(unittest.TestCase):
    def test_libreoffice_gate_rejects_changed_sources_and_files(self) -> None:
        import libreoffice_smoke as smoke

        with tempfile.TemporaryDirectory() as temp, redirect_stdout(StringIO()):
            root = Path(temp) / "fixtures"
            smoke.prepare(root)
            entries = []
            for name in smoke.NAMES:
                saved = root / (Path(name).stem + ".saved.xlsm")
                shutil.copyfile(root / name, saved)
                entries.append({"name": name, "ok": True, "saved_sha256": smoke.sha(saved)})
            report = {"ok": True, "fixtures": entries}
            smoke.write_json(root / "runtime.json", report)
            smoke.verify(root)  # Synthetic runtime report; checks only the static gate here.

            saved = root / "generated.saved.xlsm"
            baseline = saved.read_bytes()
            saved.write_bytes(baseline + b"changed")
            with self.assertRaisesRegex(RuntimeError, "changed after runtime"):
                smoke.verify(root)
            saved.write_bytes(baseline)

            original = root / smoke.NAMES[0]
            original_bytes = original.read_bytes()
            original.write_bytes(original_bytes + b"changed")
            with self.assertRaisesRegex(RuntimeError, "input changed"):
                smoke.verify(root)
            original.write_bytes(original_bytes)

            source = root / "vba" / "Runtime.bas"
            source.write_text(smoke.RUNTIME.replace(smoke.EXPECTED_TEXT, "????????"), encoding="utf-8")
            changed = root / "changed.xlsm"
            create_workbook(root / "workbook.json", changed, source=root / "vba", experimental=True, target_os="linux")
            shutil.copyfile(changed, saved)
            entries[0]["saved_sha256"] = smoke.sha(saved)
            smoke.write_json(root / "runtime.json", report)
            with self.assertRaisesRegex(RuntimeError, "VBA source changed"):
                smoke.verify(root)
            self.assertFalse(json.loads((root / "verification.json").read_text(encoding="utf-8"))["ok"])

    def test_target_os_does_not_spoof_host(self) -> None:
        for platform, expected in [("darwin", "macos"), ("linux", "linux")]:
            with patch("sys.platform", platform):
                report = capabilities("windows")
                self.assertEqual(report["host_os"], expected)
                self.assertEqual(report["target_os"], "windows")
                self.assertFalse(report["excel_worker_available"])
                with patch("usable_xlsm.runner.subprocess.run") as start:
                    self.assertEqual(excel_pids(), set())
                    self.assertEqual(run_job("run", "book.xlsm").error_code, "excel_backend_unavailable")
                    start.assert_not_called()

    def test_doctor_static_succeeds_without_excel(self) -> None:
        for platform in ["linux", "darwin"]:
            with patch("sys.platform", platform):
                result = CliRunner().invoke(app, ["doctor"])
                self.assertEqual(result.exit_code, 0, result.output)
                self.assertEqual(json.loads(result.output)["mode"], "static")
                self.assertFalse(json.loads(result.output)["runtime_verified"])

    def test_doctor_windows_auto_still_requires_excel(self) -> None:
        registry = SimpleNamespace(HKEY_CURRENT_USER=1, OpenKey=lambda *args: (_ for _ in ()).throw(OSError()))
        client = ModuleType("win32com.client")
        package = ModuleType("win32com")
        package.client = client
        with (
            patch("sys.platform", "win32"),
            patch.dict("sys.modules", {"winreg": registry, "win32com": package, "win32com.client": client}),
            patch("usable_xlsm.cli.excel_pids", return_value=set()),
        ):
            result = CliRunner().invoke(app, ["doctor"])
        self.assertEqual(result.exit_code, 1, repr(result.exception))
        self.assertTrue(result.output, repr(result.exception))
        self.assertFalse(json.loads(result.output)["ok"])

    def test_single_file_syntax_error_returns_json(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            source = Path(raw) / "Broken.bas"
            source.write_text("Public Sub Broken()\nIf Then\nEnd Sub\n", encoding="utf-8")
            result = CliRunner().invoke(app, ["check", "--source", str(source)])
        self.assertEqual(result.exit_code, 1, result.output)
        report = json.loads(result.output)
        self.assertFalse(report["ok"])
        self.assertEqual(report["issues"][0]["file"], "Broken.bas")

    def test_portability_ignores_comments_and_strings(self) -> None:
        code = '''Public Const E As Long = vbObjectError + 1
Public Sub Work()
    x = "FormatConditions: vbObjectError"
    ' FormatConditions: Declare Lib
    Rem CreateObject: vbObjectError
    obj.FormatConditions.Add 1
    Case "x": Work
Label:
End Sub
'''
        issues = portability_issues({"Work.bas": code}, "libreoffice")
        self.assertEqual([item["code"] for item in issues], ["error_constant", "conditional_format", "multi_statement"])
        self.assertNotIn("error_constant", [item["code"] for item in portability_issues({"Work.bas": code}, "windows")])

    def test_assertion_constant_preserves_error_number_and_runner_has_no_colon_statements(self) -> None:
        self.assertIn(str(-2147221504 + 9001), assert_module_source())
        source = generate_runner(discover({"Tests.bas": "Public Sub Test_One()\nEnd Sub\n"}))
        self.assertEqual(portability_issues({"Runner.bas": source}, "libreoffice"), [])

    def test_partial_validation_preserves_omitted_modules_and_rejects_additions(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            (root / "Main.bas").write_text("Public Sub Main()\nEnd Sub\n", encoding="utf-8")
            with patch("usable_xlsm.core.extract_vba", return_value={"Main.bas": "", "Other.bas": ""}):
                self.assertEqual(set(validate_vba_modules(root, "book.xlsm", partial=True)), {"Main.bas"})
                with self.assertRaises(VbaModuleMismatchError):
                    validate_vba_modules(root, "book.xlsm")
                with self.assertRaises(ValueError):
                    validate_vba_modules(root, "book.xlsm", partial=True, sync=True)
            with patch("usable_xlsm.core.extract_vba", return_value={"Other.bas": ""}):
                with self.assertRaises(VbaModuleMismatchError):
                    validate_vba_modules(root, "book.xlsm", partial=True)

    def test_plan_partial_includes_unchanged_tests_without_deletion(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            (root / "Main.bas").write_text("Public Sub Main()\n    x = 2\nEnd Sub\n", encoding="utf-8")
            modules = {"Main.bas": "Public Sub Main()\n    x = 1\nEnd Sub\n", "Tests.bas": "Public Sub Test_One()\nEnd Sub\n"}
            with patch("usable_xlsm.planning.extract_vba", return_value=modules):
                report = plan_changes(root, "book.xlsm", partial=True)
        self.assertEqual(report["changed"], ["Main.bas"])
        self.assertEqual(report["removed"], [])
        self.assertEqual(report["tests"], ["Tests.Test_One"])
        self.assertNotIn("diff", report)

    def test_vba_compression_round_trip_and_chunk_boundaries(self) -> None:
        for data in [b"", b"A", b"AB" * 4000, bytes(range(256)) * 50, random.Random(1).randbytes(8192)]:
            encoded = compress_vba(data)
            self.assertEqual(decompress_stream(bytearray(encoded)), data)
            position, decoded_sizes = 1, []
            while position < len(encoded):
                header = struct.unpack_from("<H", encoded, position)[0]
                size = (header & 0xFFF) + 3
                chunk = encoded[position:position + size]
                decoded_sizes.append(len(decompress_stream(bytearray(b"\x01" + chunk))))
                position += size
            self.assertTrue(all(size == 4096 for size in decoded_sizes[:-1]))
        with self.assertRaises(ValueError):
            compress_vba(random.Random(1).randbytes(4000))

    def test_project_properties_match_ms_ovba_example(self) -> None:
        identifier = "{917DED54-440B-4FD1-A5C1-74ACF261E600}"
        self.assertEqual(_project_property(bytes(4), identifier, 7), "0705D8E3D8EDDBF1DBF1DBF1DBF1")
        self.assertEqual(_project_property(bytes(1), identifier, 14), "0E0CD1ECDFF4E7F5E7F5E7")
        self.assertEqual(_project_property(bytes([255]), identifier, 21), "1517CAF1D6F9D7F9D706")

    def test_cfb_mini_and_regular_streams_and_directory_trees(self) -> None:
        for count in range(1, 32):
            streams = {f"VBA/Module{i}": bytes([i]) * (i * 501 + 1) for i in range(count)}
            streams["PROJECT"] = b"project"
            with OleFileIO(BytesIO(_compound_file(streams)), raise_defects=40) as ole:
                self.assertEqual(set("/".join(path) for path in ole.listdir()), set(streams))
                for path, data in streams.items():
                    self.assertEqual(ole.openstream(path).read(), data)

                def black_height(index: int) -> int:
                    if index == 0xFFFFFFFF:
                        return 1
                    entry = ole.direntries[index]
                    left, right = black_height(entry.sid_left), black_height(entry.sid_right)
                    self.assertEqual(left, right)
                    if entry.color == 0:
                        for child in [entry.sid_left, entry.sid_right]:
                            if child != 0xFFFFFFFF:
                                self.assertEqual(ole.direntries[child].color, 1)
                    return left + entry.color

                for parent in [ole.root, ole.direntries[1]]:
                    self.assertEqual(ole.direntries[parent.sid_child].color, 1)
                    black_height(parent.sid_child)

    def test_new_workbook_round_trip_buttons_and_no_overwrite(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            source = root / "src"
            source.mkdir()
            text = 'Public Sub Hello()\n    MsgBox "こんにちは"\nEnd Sub\n'
            (source / "Hello.bas").write_text(text, encoding="utf-8")
            spec = root / "spec.json"
            spec.write_text(json.dumps({"sheets": [{"name": "入力", "cells": {"A1": "=literal"}, "buttons": [{"cell": "B2", "macro": "Hello.Hello", "caption": "実行"}]}]}, ensure_ascii=False), encoding="utf-8")
            output = root / "new.xlsm"
            report = create_workbook(spec, output, source=source, experimental=True)
            self.assertFalse(report["runtime_verified"])
            modules = extract_vba(output)
            self.assertIn('MsgBox "こんにちは"', modules["Hello.bas"])
            with ZipFile(output) as archive:
                self.assertIn(b"macroEnabled", archive.read("[Content_Types].xml"))
                self.assertNotIn(b"<f>", archive.read("xl/worksheets/sheet1.xml"))
                self.assertIn(b"Hello.Hello", archive.read("xl/drawings/vmlDrawing1.vml"))
            before = output.read_bytes()
            with self.assertRaises(FileExistsError):
                create_workbook(spec, output, source=source, experimental=True)
            self.assertEqual(output.read_bytes(), before)
            project = root / "project.bin"
            extract_project(output, project)
            other = root / "seed-copy.xlsm"
            create_workbook(spec, other, vba_project=project)
            self.assertEqual(extract_vba(other), modules)

    def test_builder_rejects_classes_unencodable_text_and_name_mismatch(self) -> None:
        for name, text, error in [("Class.cls", "", ValueError), ("Main.bas", 'Attribute VB_Name = "Wrong"\nPublic Sub Main()\nEnd Sub\n', ValueError), ("Main.bas", "Public Sub Main()\n    '😀\nEnd Sub\n", UnicodeEncodeError)]:
            with tempfile.TemporaryDirectory() as raw:
                root = Path(raw)
                (root / name).write_text(text, encoding="utf-8")
                with self.assertRaises(error):
                    project_bytes(root)


if __name__ == "__main__":
    unittest.main()
