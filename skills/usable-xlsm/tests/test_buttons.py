from __future__ import annotations

import unittest

from usable_xlsm.core import normalize_button_specs, normalize_cell_specs, normalize_worksheet_specs


class ButtonSpecTests(unittest.TestCase):
    def test_normalizes_worksheet_and_cell_specs(self) -> None:
        self.assertEqual(
            normalize_worksheet_specs([{"name": "Calculator", "create": True, "visible": "visible"}]),
            [{"name": "Calculator", "create": True, "visible": "visible"}],
        )
        self.assertEqual(
            normalize_cell_specs([{"sheet": "Calculator", "address": "b2", "value": 0}]),
            [{
                "sheet": "Calculator",
                "address": "B2",
                "value": 0,
                "formula": None,
                "number_format": None,
                "overwrite": True,
            }],
        )

    def test_rejects_ambiguous_cell_and_worksheet_specs(self) -> None:
        with self.assertRaisesRegex(ValueError, "exactly one"):
            normalize_cell_specs([{"sheet": "Sheet1", "address": "A1", "value": 1, "formula": "=1"}])
        with self.assertRaisesRegex(ValueError, "Duplicate worksheet"):
            normalize_worksheet_specs([{"name": "Sheet1"}, {"name": "sheet1"}])

    def test_normalizes_defaults_and_manifest_wrapper(self) -> None:
        result = normalize_button_specs(
            {
                "buttons": [
                    {"sheet": "Sheet1", "name": "btnRun", "macro": "Module1.Run"}
                ]
            }
        )
        self.assertEqual(
            result,
            [
                {
                    "sheet": "Sheet1",
                    "name": "btnRun",
                    "caption": "btnRun",
                    "macro": "Module1.Run",
                    "left": 10.0,
                    "top": 10.0,
                    "width": 120.0,
                    "height": 24.0,
                    "replace": False,
                }
            ],
        )

    def test_rejects_duplicate_names_and_invalid_geometry(self) -> None:
        with self.assertRaisesRegex(ValueError, "Duplicate"):
            normalize_button_specs(
                [
                    {"sheet": "Sheet1", "name": "btn", "macro": "M.Run"},
                    {"sheet": "Sheet1", "name": "btn", "macro": "M.Run"},
                ]
            )
        with self.assertRaisesRegex(ValueError, "finite"):
            normalize_button_specs(
                [{"sheet": "Sheet1", "name": "btn", "macro": "M.Run", "width": -1}]
            )


if __name__ == "__main__":
    unittest.main()
