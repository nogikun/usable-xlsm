from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest.mock import patch

import tomllib

from usable_xlsm.runner import excel_backend_error, run_job


class PlatformBackendTests(unittest.TestCase):
    def test_pywin32_is_windows_only_dependency(self) -> None:
        pyproject = Path(__file__).parents[1] / "pyproject.toml"
        project = tomllib.loads(pyproject.read_text(encoding="utf-8"))["project"]
        pywin32 = next(item for item in project["dependencies"] if item.startswith("pywin32"))
        self.assertIn("sys_platform == 'win32'", pywin32)

    def test_non_windows_backend_fails_before_starting_excel(self) -> None:
        with patch.object(sys, "platform", "darwin"):
            self.assertIsNotNone(excel_backend_error())
            result = run_job("run_macro", "book.xlsm")
        self.assertFalse(result.ok)
        self.assertEqual(result.error_code, "excel_backend_unavailable")
        self.assertIn("Windows", result.error or "")


if __name__ == "__main__":
    unittest.main()
