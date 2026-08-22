from __future__ import annotations

import shutil
import tempfile
from pathlib import Path
import unittest
from unittest.mock import patch

from usable_xlsm.core import TestRun, VbaUpdateError, restore_workbook, update_vba
from usable_xlsm.runner import JobResult
from usable_xlsm.security import PreflightReport, sha256_file
from usable_xlsm.testing import TestResult


class AtomicUpdateTests(unittest.TestCase):
    def authorized(self, workbook: Path) -> PreflightReport:
        return PreflightReport(
            workbook=workbook,
            operation="edit",
            sha256=sha256_file(workbook),
            size=workbook.stat().st_size,
            trusted=True,
            allowed=True,
            trust_source="operator",
            has_vba=True,
        )

    def test_failed_worker_never_changes_original(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            workbook = root / "book.xlsm"
            workbook.write_bytes(b"original")
            source = root / "src"
            source.mkdir()
            (source / "Module1.bas").write_text("Sub X()\nEnd Sub", encoding="utf-8")
            before = workbook.read_bytes()
            with (
                patch("usable_xlsm.core.preflight_workbook", return_value=self.authorized(workbook)),
                patch("usable_xlsm.core.validate_vba_modules", return_value={"Module1.bas": "Sub X()\nEnd Sub"}),
                patch("usable_xlsm.core.extract_vba", return_value={"Module1.bas": "Sub X()\nEnd Sub"}),
                patch("usable_xlsm.core.check_directory", return_value=[]),
                patch(
                    "usable_xlsm.core.run_job",
                    return_value=JobResult(
                        ok=False,
                        error="boom",
                        error_code="excel_host_not_clean",
                    ),
                ),
            ):
                with self.assertRaises(VbaUpdateError) as raised:
                    update_vba(source, workbook, trust_workbook=True, run_test_suite=False)
            self.assertEqual(raised.exception.error_code, "excel_host_not_clean")
            self.assertEqual(workbook.read_bytes(), before)
            self.assertFalse(list(root.glob("*.staging.xlsm")))

    def test_success_promotes_staging_and_records_hashes(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            workbook = root / "book.xlsm"
            workbook.write_bytes(b"original")
            source = root / "src"
            source.mkdir()
            (source / "Module1.bas").write_text("Sub X()\nEnd Sub", encoding="utf-8")

            def update_staging(action: str, target: Path, **_: object) -> JobResult:
                Path(target).write_bytes(b"updated")
                return JobResult(
                    ok=True,
                    data={
                        "updated": ["Module1"],
                        "added": [],
                        "removed": [],
                        "components": [{"name": "Module1", "type": 1}],
                    },
                )

            with (
                patch("usable_xlsm.core.preflight_workbook", return_value=self.authorized(workbook)),
                patch("usable_xlsm.core.validate_vba_modules", return_value={"Module1.bas": "Sub X()\nEnd Sub"}),
                patch("usable_xlsm.core.extract_vba", return_value={"Module1.bas": "Sub X()\nEnd Sub"}),
                patch("usable_xlsm.core.check_directory", return_value=[]),
                patch("usable_xlsm.core.run_job", side_effect=update_staging),
                patch("usable_xlsm.core.verify_applied_modules"),
                patch(
                    "usable_xlsm.core.run_tests",
                    return_value=TestRun(
                        ok=True,
                        results=[TestResult("Tests", "Test_X", "pass", 0.1, "")],
                    ),
                ),
            ):
                report = update_vba(source, workbook, trust_workbook=True)
            self.assertEqual(workbook.read_bytes(), b"updated")
            self.assertNotEqual(report["sha256_before"], report["sha256_after"])
            self.assertTrue(Path(report["backup"]).is_file())
            self.assertTrue(Path(report["audit_log"]).is_file())
            audit = Path(report["audit_log"]).read_text(encoding="utf-8")
            self.assertIn('"status": "ready_to_promote"', audit)
            self.assertIn('"status": "succeeded"', audit)

    def test_restore_is_atomic_and_audited(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            workbook = root / "book.xlsm"
            backup = root / "book.bak.xlsm"
            workbook.write_bytes(b"current")
            backup.write_bytes(b"restored")
            report = restore_workbook(workbook, backup_path=backup)
            self.assertEqual(workbook.read_bytes(), b"restored")
            audit = Path(report["audit_log"]).read_text(encoding="utf-8")
            self.assertIn('"status": "ready_to_restore"', audit)
            self.assertIn('"status": "succeeded"', audit)


if __name__ == "__main__":
    unittest.main()
