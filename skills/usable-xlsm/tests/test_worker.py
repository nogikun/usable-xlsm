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


class FakeButton:
    def __init__(self, name: str = "") -> None:
        self.Name = name
        self.Caption = ""
        self.OnAction = ""
        self.deleted = False

    def Delete(self) -> None:
        self.deleted = True


class FakeRange:
    def __init__(self) -> None:
        self.Value = None
        self.Formula = ""
        self.NumberFormat = "General"


class FakeButtons:
    def __init__(self) -> None:
        self.items: dict[str, FakeButton] = {}
        self.added: list[tuple[float, float, float, float]] = []

    def Item(self, name: str) -> FakeButton:
        if name in self.items:
            return self.items[name]
        for button in self.items.values():
            if button.Name == name:
                return button
        raise KeyError(name)

    def Add(self, left: float, top: float, width: float, height: float) -> FakeButton:
        self.added.append((left, top, width, height))
        button = FakeButton()
        button.Left = left
        button.Top = top
        button.Width = width
        button.Height = height
        self.items[f"Button {len(self.items) + 1}"] = button
        return button


class FakeSheet:
    def __init__(self, name: str = "Sheet1") -> None:
        self.Name = name
        self.Visible = -1
        self.buttons = FakeButtons()
        self.ranges: dict[str, FakeRange] = {}

    def Buttons(self) -> FakeButtons:
        return self.buttons

    def Range(self, address: str) -> FakeRange:
        return self.ranges.setdefault(address, FakeRange())


class FakeWorksheets:
    def __init__(self, *sheets: FakeSheet) -> None:
        self.items = {sheet.Name: sheet for sheet in sheets}

    @property
    def Count(self) -> int:
        return len(self.items)

    def Item(self, key: object) -> FakeSheet:
        if isinstance(key, int):
            return list(self.items.values())[key - 1]
        if str(key) in self.items:
            return self.items[str(key)]
        for sheet in self.items.values():
            if sheet.Name == str(key):
                return sheet
        raise KeyError(key)

    def __call__(self, key: object) -> FakeSheet:
        return self.Item(key)

    def Add(self, **_: object) -> FakeSheet:
        sheet = FakeSheet(f"Sheet{len(self.items) + 1}")
        self.items[sheet.Name] = sheet
        return sheet


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

    def test_update_places_declared_form_control_button(self) -> None:
        sheet = FakeSheet()
        book = SimpleNamespace(
            Name="book.xlsm",
            VBProject=SimpleNamespace(VBComponents=[]),
            Worksheets=lambda name: sheet if name == "Sheet1" else None,
            Save=Mock(),
        )

        report = _job_update(
            Mock(),
            book,
            {
                "modules": {},
                "buttons": [
                    {
                        "sheet": "Sheet1",
                        "name": "btnRun",
                        "caption": "Run",
                        "macro": "Module1.Run",
                        "left": 12.0,
                        "top": 24.0,
                        "width": 100.0,
                        "height": 22.0,
                        "replace": False,
                    }
                ],
            },
        )

        button = next(iter(sheet.buttons.items.values()))
        self.assertEqual(sheet.buttons.added, [(12.0, 24.0, 100.0, 22.0)])
        self.assertEqual(button.Name, "btnRun")
        self.assertEqual(button.Caption, "Run")
        self.assertEqual(button.OnAction, "Module1.Run")
        self.assertEqual(report["buttons"][0]["name"], "btnRun")
        self.assertEqual(report["buttons_verified"], [{"sheet": "Sheet1", "name": "btnRun", "verified": True}])
        book.Save.assert_called_once_with()

    def test_update_creates_worksheet_and_writes_cells(self) -> None:
        existing = FakeSheet("Sheet1")
        worksheets = FakeWorksheets(existing)
        book = SimpleNamespace(
            Name="book.xlsm",
            VBProject=SimpleNamespace(VBComponents=[]),
            Worksheets=worksheets,
            Save=Mock(),
        )
        report = _job_update(
            Mock(),
            book,
            {
                "modules": {},
                "worksheets": [{"name": "Calculator", "create": True, "visible": "visible"}],
                "cells": [
                    {
                        "sheet": "Calculator",
                        "address": "B2",
                        "value": 0,
                        "formula": None,
                        "number_format": "0.00",
                        "overwrite": True,
                    },
                    {
                        "sheet": "Calculator",
                        "address": "C2",
                        "value": None,
                        "formula": "=B2+1",
                        "number_format": None,
                        "overwrite": True,
                    },
                ],
            },
        )
        calculator = worksheets.Item("Calculator")
        self.assertEqual(calculator.Range("B2").Value, 0)
        self.assertEqual(calculator.Range("B2").NumberFormat, "0.00")
        self.assertEqual(calculator.Range("C2").Formula, "=B2+1")
        self.assertEqual(report["worksheets_verified"], [{"name": "Calculator", "verified": True}])
        self.assertEqual(
            report["cells_verified"],
            [
                {"sheet": "Calculator", "address": "B2", "verified": True},
                {"sheet": "Calculator", "address": "C2", "verified": True},
            ],
        )
        book.Save.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
