from __future__ import annotations

from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import zipfile

from usable_xlsm.core import apply_script, verify_saved_buttons, verify_saved_workbook_objects


class ApplyScriptTests(unittest.TestCase):
    def test_saved_button_verification_reads_closed_package(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            workbook = Path(raw) / "book.xlsm"
            files = {
                "xl/workbook.xml": (
                    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
                ),
                "xl/_rels/workbook.xml.rels": (
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId1" Target="worksheets/sheet1.xml" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>'
                    '</Relationships>'
                ),
                "xl/worksheets/_rels/sheet1.xml.rels": (
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId1" Target="../drawings/vmlDrawing1.vml" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"/>'
                    '</Relationships>'
                ),
                "xl/worksheets/sheet1.xml": (
                    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                    '<sheetData><row r="1">'
                    '<c r="A1" t="inlineStr"><is><t>Hello</t></is></c>'
                    '<c r="B1"><v>12</v></c>'
                    '<c r="C1"><f>B1+1</f><v>13</v></c>'
                    '</row></sheetData></worksheet>'
                ),
                "xl/drawings/vmlDrawing1.vml": (
                    '<xml xmlns:v="urn:schemas-microsoft-com:vml" '
                    'xmlns:x="urn:schemas-microsoft-com:office:excel">'
                    '<v:shape id="btnRun" style="position:absolute;margin-left:12pt;'
                    'margin-top:24pt;width:100pt;height:22pt">'
                    '<v:textbox><div>Run</div></v:textbox>'
                    '<x:ClientData ObjectType="Button"><x:FmlaMacro>[0]!Module1.Run'
                    '</x:FmlaMacro></x:ClientData></v:shape></xml>'
                ),
            }
            with zipfile.ZipFile(workbook, "w") as package:
                for name, content in files.items():
                    package.writestr(name, content)
            result = verify_saved_buttons(
                workbook,
                [{
                    "sheet": "Sheet1",
                    "name": "btnRun",
                    "caption": "Run",
                    "macro": "Module1.Run",
                    "left": 12,
                    "top": 24,
                    "width": 100,
                    "height": 22,
                }],
            )
            self.assertEqual(result, [{"sheet": "Sheet1", "name": "btnRun", "verified": True}])
            self.assertEqual(
                verify_saved_workbook_objects(
                    workbook,
                    [{"name": "Sheet1", "create": False, "visible": "visible"}],
                    [
                        {
                            "sheet": "Sheet1",
                            "address": "A1",
                            "value": "Hello",
                            "formula": None,
                            "number_format": None,
                            "overwrite": True,
                        },
                        {
                            "sheet": "Sheet1",
                            "address": "B1",
                            "value": 12,
                            "formula": None,
                            "number_format": None,
                            "overwrite": True,
                        },
                        {
                            "sheet": "Sheet1",
                            "address": "C1",
                            "value": None,
                            "formula": "=B1+1",
                            "number_format": None,
                            "overwrite": True,
                        },
                    ],
                ),
                {
                    "worksheets": [{"name": "Sheet1", "verified": True}],
                    "cells": [
                        {"sheet": "Sheet1", "address": "A1", "verified": True},
                        {"sheet": "Sheet1", "address": "B1", "verified": True},
                        {"sheet": "Sheet1", "address": "C1", "verified": True},
                    ],
                },
            )

    def test_script_rejects_unsupported_keys_instead_of_ignoring_them(self) -> None:
        with self.assertRaisesRegex(ValueError, "Unsupported apply script key"):
            apply_script({"not_a_supported_operation": []}, "book.xlsm")

    def test_script_merges_modules_and_buttons_into_one_update(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            workbook = root / "book.xlsm"
            workbook.write_bytes(b"workbook")
            module = root / "Module1.bas"
            module.write_text("Public Sub Run()\nEnd Sub\n", encoding="utf-8")
            captured: dict[str, object] = {}

            def fake_update(source: Path, target: Path, **kwargs: object) -> dict[str, object]:
                captured["source"] = source
                captured["target"] = target
                captured.update(kwargs)
                names = sorted(path.name for path in Path(source).iterdir())
                captured["source_names"] = names
                return {"ok": True, "buttons": kwargs["button_specs"]}

            with (
                patch(
                    "usable_xlsm.core.extract_vba",
                    return_value={
                        "Module1.bas": "Public Sub Old()\nEnd Sub\n",
                        "Sheet1.cls": "Attribute VB_Name = \"Sheet1\"\n",
                    },
                ),
                patch("usable_xlsm.core.update_vba", side_effect=fake_update),
            ):
                report = apply_script(
                    {
                        "modules": [module.name],
                        "buttons": [
                            {"sheet": "Sheet1", "name": "btnRun", "macro": "Module1.Run"}
                        ],
                    },
                    workbook,
                    script_dir=root,
                )

            self.assertEqual(captured["source_names"], ["Module1.bas", "Sheet1.cls"])
            self.assertTrue(captured["sync"])
            self.assertEqual(captured["target"], workbook)
            self.assertEqual(report["script_modules"], ["Module1.bas"])


if __name__ == "__main__":
    unittest.main()
