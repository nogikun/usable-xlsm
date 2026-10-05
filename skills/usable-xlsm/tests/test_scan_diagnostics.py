from __future__ import annotations

from contextlib import contextmanager
import json
from pathlib import Path
import subprocess
import tempfile
import time
import unittest
from unittest.mock import patch
import zipfile

from typer.testing import CliRunner

from usable_xlsm.cli import app
from usable_xlsm.core import run_macro, run_tests, update_vba
from usable_xlsm.security import WorkbookSecurityError, preflight_workbook


@contextmanager
def inject_scanner(code: str):
    """Inject only at the parser call inside the actual disposable worker."""
    original_run = subprocess.run

    def run(command, **kwargs):
        command = list(command)
        # Injection runs after the production bootstrap restores sys.path, so
        # this also exercises the real interpreter on Windows, not its launcher.
        assert "scan(request)" in command[-1]
        command[-1] = command[-1].replace("scan(request)", f"exec({code!r}); scan(request)")
        return original_run(command, **kwargs)

    with patch("usable_xlsm.security.subprocess.run", side_effect=run):
        yield


class ScanDiagnosticsTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.book = self.root / "book.xlsm"
        with zipfile.ZipFile(self.book, "w") as package:
            package.writestr("[Content_Types].xml", "<Types/>")
        self.before = self.book.read_bytes()
        self.policy = self.root / "policy.toml"
        self.policy.write_text(
            "[security]\nallow_scan_failures=true\nscan_timeout_seconds=2.0\n", encoding="utf-8",
        )

    def test_slow_deobfuscation_reports_phase_and_policy_deadline(self):
        code = (
            "import time\nfrom oletools.olevba import VBA_Parser\n"
            "def slow(self, **kwargs):\n"
            "    assert kwargs['show_decoded_strings'] is True\n"
            "    assert kwargs['deobfuscate'] is True\n"
            "    time.sleep(30)\n"
            "VBA_Parser.analyze_macros = slow\n"
        )
        started = time.monotonic()
        with inject_scanner(code):
            report = preflight_workbook(self.book, trust_workbook=True, policy_path=self.policy)
        self.assertLess(time.monotonic() - started, 8)
        self.assertTrue(report.trusted)
        self.assertFalse(report.allowed)
        self.assertEqual(report.scan_status, "timeout")
        self.assertEqual(report.scan_phase, "macro_analysis")
        self.assertEqual(report.scan_timeout_seconds, 2)
        self.assertIn("macro_scan_timeout", [f.code for f in report.findings])
        self.assertEqual(self.book.read_bytes(), self.before)

    def test_cli_timeout_override_also_bounds_parser_close(self):
        code = ("import time\nfrom oletools.olevba import VBA_Parser\n"
                "VBA_Parser.close = lambda self: time.sleep(30)\n")
        with inject_scanner(code):
            result = CliRunner().invoke(app, [
                "preflight", "--workbook", str(self.book), "--operation", "inspect",
                "--policy", str(self.policy), "--scan-timeout", "1.5",
            ])
        self.assertEqual(result.exit_code, 2, result.output)
        report = json.loads(result.output)
        self.assertFalse(report["allowed"])
        self.assertEqual(report["scan_status"], "timeout")
        self.assertEqual(report["scan_phase"], "parser_close")
        self.assertEqual(report["scan_timeout_seconds"], 1.5)
        self.assertNotIn("no_vba", [f["code"] for f in report["findings"]])

    def test_crash_reports_phase_and_cannot_be_trusted_away(self):
        code = ("from oletools.olevba import VBA_Parser\n"
                "def crash(self, **kwargs):\n    raise SystemExit(9)\n"
                "VBA_Parser.analyze_macros = crash\n")
        with inject_scanner(code):
            report = preflight_workbook(self.book, trust_workbook=True, policy_path=self.policy)
        self.assertFalse(report.allowed)
        self.assertEqual(report.scan_status, "failed")
        self.assertEqual(report.scan_phase, "parser_close")
        self.assertIn("macro_scan_worker_failed", [f.code for f in report.findings])

    def test_completed_scan_preserves_large_unicode_findings_and_diagnostics(self):
        code = ("from oletools.olevba import VBA_Parser\n"
                "def analysis(self, **kwargs):\n"
                "    print('diagnostic ' * 10000)\n"
                "    return [('Suspicious', 'Shell', '\u65e5\u672c\u8a9e' * 1000)] * 100\n"
                "VBA_Parser.analyze_macros = analysis\n")
        with inject_scanner(code):
            report = preflight_workbook(self.book, operation="inspect", scan_timeout=5)
        self.assertTrue(report.allowed)
        self.assertEqual(report.scan_status, "complete")
        self.assertEqual(report.scan_phase, "parser_close")
        warnings = [f for f in report.findings if f.code == "olevba_suspicious"]
        self.assertEqual(len(warnings), 100)
        self.assertEqual(warnings[0].detail, "\u65e5\u672c\u8a9e" * 1000)

    def test_timeout_blocks_export_execution_and_tests_before_excel(self):
        source = self.root / "src"
        source.mkdir()
        (source / "MainModule.bas").write_text("Sub X()\nEnd Sub\n", encoding="utf-8")
        scan = {
            "has_vba": True, "has_xlm": False, "stomping": False, "vba_checked": True,
            "phase": "macro_analysis", "scan_status": "timeout",
            "findings": [], "blockers": ["workbook macro scan exceeded its deadline"],
        }
        actions = (
            lambda: update_vba(source, self.book, trust_workbook=True, run_test_suite=False, skip_unchanged=True),
            lambda: run_macro(self.book, "MainModule.X", trust_workbook=True),
            lambda: run_tests(self.book, trust_workbook=True, require_tests=False),
        )
        with patch("usable_xlsm.security._scan_workbook", return_value=scan), \
             patch("usable_xlsm.core.run_job") as excel:
            for action in actions:
                with self.assertRaises(WorkbookSecurityError) as caught:
                    action()
                self.assertEqual(caught.exception.report.scan_phase, "macro_analysis")
            excel.assert_not_called()
        self.assertEqual(self.book.read_bytes(), self.before)
        self.assertEqual(sorted(p.name for p in self.root.iterdir()), ["book.xlsm", "policy.toml", "src"])

    def test_invalid_policy_and_api_deadlines_do_not_start_a_scanner(self):
        with patch("usable_xlsm.security._scan_workbook") as scanner:
            for deadline in (0, -1, float("inf"), float("nan"), True):
                with self.subTest(deadline=deadline), self.assertRaises(ValueError):
                    preflight_workbook(self.book, scan_timeout=deadline)
            for value in ("0", "-1", "inf", "nan", "true"):
                self.policy.write_text(f"[security]\nscan_timeout_seconds={value}\n", encoding="utf-8")
                with self.subTest(value=value), self.assertRaises(ValueError):
                    preflight_workbook(self.book, policy_path=self.policy)
            scanner.assert_not_called()

    def test_partial_or_unchecked_result_cannot_authorize(self):
        state = {"has_vba": False, "has_xlm": False, "stomping": False,
                 "vba_checked": False, "phase": "macro_analysis", "findings": [], "blockers": []}
        for event in (
            {"type": "progress", "state": state},
            {"type": "result", "state": {**state, "scan_status": "complete"}},
        ):
            output = json.dumps(event).encode()
            with self.subTest(event=event), patch("usable_xlsm.security.subprocess.run", return_value=
                subprocess.CompletedProcess([], 0, stdout=output, stderr=b"")):
                report = preflight_workbook(self.book, trust_workbook=True, policy_path=self.policy)
            self.assertFalse(report.allowed)
            self.assertEqual(report.scan_status, "failed")


if __name__ == "__main__":
    unittest.main()
