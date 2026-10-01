from __future__ import annotations

import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from usable_xlsm.runner import _kill, _wait_for_excel_exit


class RunnerSafetyTests(unittest.TestCase):
    def test_owned_excel_exit_wait_closes_handle_and_fails_on_timeout(self) -> None:
        class ProcessError(Exception):
            def __init__(self, code):
                self.winerror = code

        api = SimpleNamespace(OpenProcess=Mock(return_value="owned-handle"), CloseHandle=Mock())
        event = SimpleNamespace(WaitForSingleObject=Mock(side_effect=[258, 0]), WAIT_OBJECT_0=0)
        with patch.dict("sys.modules", {"win32api": api, "win32event": event,
                "win32con": SimpleNamespace(SYNCHRONIZE=0x100000),
                "pywintypes": SimpleNamespace(error=ProcessError)}):
            self.assertFalse(_wait_for_excel_exit(1234))
            self.assertTrue(_wait_for_excel_exit(1234))
            self.assertEqual(api.CloseHandle.call_count, 2)
            api.OpenProcess.assert_called_with(0x100000, False, 1234)
            event.WaitForSingleObject.assert_called_with("owned-handle", 5000)
            api.OpenProcess.side_effect = ProcessError(87)
            self.assertTrue(_wait_for_excel_exit(1234))
            api.OpenProcess.side_effect = ProcessError(5)
            self.assertFalse(_wait_for_excel_exit(1234))

    def test_worker_only_kill_does_not_include_process_tree(self) -> None:
        with patch("usable_xlsm.runner.subprocess.run") as run:
            _kill(1234, tree=False)
        command = run.call_args.args[0]
        self.assertEqual(command, ["taskkill", "/PID", "1234", "/F"])
        self.assertNotIn("/T", command)

    def test_owned_excel_kill_includes_process_tree(self) -> None:
        with patch("usable_xlsm.runner.subprocess.run") as run:
            _kill(1234, tree=True)
        self.assertIn("/T", run.call_args.args[0])


if __name__ == "__main__":
    unittest.main()
