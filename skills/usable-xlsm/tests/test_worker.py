from __future__ import annotations

from types import SimpleNamespace
import unittest
from unittest.mock import Mock

from usable_xlsm._worker import _job_update


class FakeCodeModule:
    CountOfLines = 0

    def DeleteLines(self, *_: object) -> None:
        pass

    def AddFromString(self, _: str) -> None:
        pass


class WorkerUpdateTests(unittest.TestCase):
    def test_update_runs_post_macro_before_saving_staging_copy(self) -> None:
        component = SimpleNamespace(
            Name="Module1",
            Type=1,
            CodeModule=FakeCodeModule(),
        )
        book = SimpleNamespace(
            Name="book.xlsm",
            VBProject=SimpleNamespace(VBComponents=[component]),
            Save=Mock(),
        )
        excel = Mock()

        report = _job_update(
            excel,
            book,
            {
                "modules": {"Module1.bas": "Sub X()\nEnd Sub"},
                "sync": False,
                "post_macro": "Module1.InstallControlButtons",
                "post_macro_args": [],
            },
        )

        excel.Run.assert_called_once_with("'book.xlsm'!Module1.InstallControlButtons")
        book.Save.assert_called_once_with()
        self.assertEqual(report["post_macro"], "Module1.InstallControlButtons")


if __name__ == "__main__":
    unittest.main()
