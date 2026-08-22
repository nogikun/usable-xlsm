from __future__ import annotations

import unittest
from unittest.mock import patch

from usable_xlsm.runner import _kill


class RunnerSafetyTests(unittest.TestCase):
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
