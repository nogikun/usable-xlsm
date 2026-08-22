from __future__ import annotations

import json
import tempfile
from pathlib import Path
import unittest
from unittest.mock import patch

from usable_xlsm.audit import write_audit_event
from usable_xlsm.locking import FileLock, LockBusyError
from usable_xlsm.support import create_support_bundle


class OperationSafetyTests(unittest.TestCase):
    def test_active_lock_is_not_stolen(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            lock_path = Path(raw) / "job.lock"
            first = FileLock(lock_path, purpose="test", job_id="one").acquire()
            try:
                with patch("usable_xlsm.locking._pid_is_running", return_value=True):
                    with self.assertRaises(LockBusyError):
                        FileLock(lock_path, purpose="test", job_id="two").acquire()
            finally:
                first.release()

    def test_stale_lock_is_recovered(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            lock_path = Path(raw) / "job.lock"
            lock_path.write_text('{"pid": 999999, "job_id": "old"}', encoding="utf-8")
            with patch("usable_xlsm.locking._pid_is_running", return_value=False):
                with FileLock(lock_path, purpose="test", job_id="new"):
                    payload = json.loads(lock_path.read_text(encoding="utf-8"))
                    self.assertEqual(payload["job_id"], "new")
            self.assertFalse(lock_path.exists())

    def test_audit_omits_full_path_and_payload_values(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            workbook = root / "secret" / "finance.xlsm"
            destination = root / "audit.jsonl"
            write_audit_event(
                workbook,
                job_id="job",
                action="run_macro",
                status="succeeded",
                details={"macro": "Module1.Run"},
                audit_log=destination,
            )
            text = destination.read_text(encoding="utf-8")
        self.assertIn('"workbook": "finance.xlsm"', text)
        self.assertNotIn("secret", text)

    def test_support_bundle_contains_no_workbook_bytes_or_full_path(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            workbook = root / "classified" / "finance.xlsm"
            workbook.parent.mkdir()
            workbook.write_bytes(b"TOP-SECRET-WORKBOOK-CONTENT")
            output = root / "support.json"
            with patch("usable_xlsm.support.excel_pids", return_value=set()):
                create_support_bundle(workbook, output)
            text = output.read_text(encoding="utf-8")
        self.assertNotIn("TOP-SECRET", text)
        self.assertNotIn("classified", text)
        self.assertIn('"contains_workbook_data": false', text)


if __name__ == "__main__":
    unittest.main()
