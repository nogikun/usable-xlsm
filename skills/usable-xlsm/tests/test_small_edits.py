from pathlib import Path
from types import SimpleNamespace
import json
import tempfile
import unittest
from unittest.mock import Mock, patch

from typer.testing import CliRunner

from usable_xlsm._worker import _apply_modules
from usable_xlsm.cli import app
from usable_xlsm.core import TestRun, VbaUpdateError, update_vba, validate_vba_modules
from usable_xlsm.runner import JobResult
from usable_xlsm.security import PreflightReport, sha256_file
from usable_xlsm.testing import TestResult


OLD = "Option Explicit\nPublic Sub Test_Existing()\n    AssertEqual 42, 42\nEnd Sub\n"
NEW = 'Attribute VB_Name = "NewModule"\nOption Explicit\nPublic Function Answer() As Long\n    Answer = 42\nEnd Function\n'


def authorized(book):
    return PreflightReport(workbook=book, operation="edit", sha256=sha256_file(book),
        size=book.stat().st_size, trusted=True, allowed=True, trust_source="operator", has_vba=True)


class SmallEditsTests(unittest.TestCase):
    def test_add_only_cli_plan_and_rejected_inputs(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            book, source = root / "book.xlsm", root / "NewModule.bas"
            book.write_bytes(b"original")
            source.write_text(NEW, encoding="utf-8")
            with (
                patch("usable_xlsm.core.extract_vba", return_value={"Old.bas": OLD}),
                patch("usable_xlsm.planning.extract_vba", return_value={"Old.bas": OLD}),
                patch("usable_xlsm.core.run_job") as start,
            ):
                result = CliRunner().invoke(app, ["plan", "--add-only", "--source", str(source), "--workbook", str(book), "--diff"])
                self.assertEqual(result.exit_code, 0, repr(result.exception))
                report = json.loads(result.output)
                self.assertEqual(report["added"], ["NewModule.bas"])
                self.assertEqual(report["removed"], [])
                self.assertFalse(report["requires_sync"])
                self.assertIn("Answer = 42", report["diff"]["NewModule.bas"])
                for kwargs in ({"partial": True}, {"sync": True}):
                    with self.assertRaises(ValueError):
                        validate_vba_modules(source, book, add_only=True, **kwargs)
                source.write_text(NEW.replace('"NewModule"', '"WrongName"'), encoding="utf-8")
                with self.assertRaisesRegex(ValueError, "VB_Name"):
                    validate_vba_modules(source, book, add_only=True)
                source = root / "NewModule.cls"
                source.write_text(NEW, encoding="utf-8")
                with self.assertRaisesRegex(ValueError, "standard .bas"):
                    validate_vba_modules(source, book, add_only=True)
                start.assert_not_called()
            source = root / "old.bas"
            source.write_text(OLD, encoding="utf-8")
            with patch("usable_xlsm.core.extract_vba", return_value={"Old.cls": OLD}):
                with self.assertRaisesRegex(ValueError, "already exists"):
                    validate_vba_modules(source, book, add_only=True)

    def test_worker_addition_does_not_rewrite_or_delete_other_modules(self):
        old = SimpleNamespace(Name="Old", Type=1, CodeModule=Mock(CountOfLines=2))

        class Components(list):
            Remove = Mock()

            def Add(self, kind):
                item = SimpleNamespace(Name="", Type=kind, CodeModule=Mock(CountOfLines=0))
                self.append(item)
                return item

        components = Components([old])
        project = SimpleNamespace(VBComponents=components)
        report = _apply_modules(project, {"NewModule.bas": NEW}, False, add_only=True)
        self.assertEqual(report, {"updated": [], "added": ["NewModule"], "removed": []})
        old.CodeModule.DeleteLines.assert_not_called()
        old.CodeModule.AddFromString.assert_not_called()
        components.Remove.assert_not_called()
        with self.assertRaisesRegex(ValueError, "already exists"):
            _apply_modules(project, {"newmodule.bas": NEW}, False, add_only=True)

    def test_addition_verifies_untouched_sources_before_promotion_and_runs_tests(self):
        for corrupt_other in (False, True):
            with self.subTest(corrupt_other=corrupt_other), tempfile.TemporaryDirectory() as temp:
                root = Path(temp)
                book, source = root / "book.xlsm", root / "NewModule.bas"
                book.write_bytes(b"original")
                source.write_text(NEW, encoding="utf-8")
                current = {"Old.bas": OLD}
                staged = {**current, "NewModule.bas": NEW}
                if corrupt_other:
                    staged["Old.bas"] = OLD.replace("42, 42", "42, 99")

                def worker(action, target, **kwargs):
                    self.assertEqual(kwargs["modules"], {"NewModule.bas": NEW})
                    self.assertFalse(kwargs["sync"])
                    self.assertTrue(kwargs["add_only"])
                    Path(target).write_bytes(b"staged")
                    return JobResult(ok=True, data={"added": ["NewModule"], "components": [
                        {"name": "Old", "type": 1}, {"name": "NewModule", "type": 1}]})

                with (
                    patch("usable_xlsm.core.preflight_workbook", return_value=authorized(book)),
                    patch("usable_xlsm.core.extract_vba", side_effect=lambda path: current if Path(path) == book else staged),
                    patch("usable_xlsm.core.run_job", side_effect=worker),
                    patch("usable_xlsm.core.run_tests", return_value=TestRun(ok=True,
                        results=[TestResult("Old", "Test_Existing", "pass", 0.1, "")])) as tests,
                ):
                    if corrupt_other:
                        with self.assertRaisesRegex(VbaUpdateError, "Old.bas"):
                            update_vba(source, book, add_only=True, trust_workbook=True)
                        self.assertEqual(book.read_bytes(), b"original")
                        tests.assert_not_called()
                    else:
                        report = update_vba(source, book, add_only=True, trust_workbook=True)
                        self.assertEqual(report["changes"]["added"], ["NewModule.bas"])
                        self.assertEqual(report["changes"]["removed"], [])
                        self.assertEqual(book.read_bytes(), b"staged")
                        self.assertTrue(tests.call_args.kwargs["require_tests"])
                        self.assertTrue(tests.call_args.kwargs["isolate"])

    def test_noop_skips_excel_tests_and_backups_but_keeps_preflight_and_audit(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            book, source = root / "book.xlsm", root / "source"
            book.write_bytes(b"original")
            source.mkdir()
            (source / "Old.bas").write_text(OLD.replace("\n", "\r\n"), encoding="utf-8", newline="")
            with (
                patch("usable_xlsm.core.preflight_workbook", return_value=authorized(book)) as preflight,
                patch("usable_xlsm.core.extract_vba", return_value={"Old.bas": OLD, "Other.bas": NEW}),
                patch("usable_xlsm.core.run_job") as worker,
                patch("usable_xlsm.core.run_tests") as tests,
                patch("usable_xlsm.core.backup_workbook") as backup,
            ):
                report = update_vba(source, book, partial=True, skip_unchanged=True, trust_workbook=True)
            self.assertEqual(report["status"], "unchanged")
            self.assertEqual(report["excel_jobs"], 0)
            self.assertFalse(report["runtime_verified"])
            self.assertIsNone(report["tests"])
            self.assertEqual(report["sha256_before"], report["sha256_after"])
            self.assertEqual(book.read_bytes(), b"original")
            self.assertIn('"status": "unchanged"', Path(report["audit_log"]).read_text(encoding="utf-8"))
            preflight.assert_called_once()
            worker.assert_not_called()
            tests.assert_not_called()
            backup.assert_not_called()


if __name__ == "__main__":
    unittest.main()
