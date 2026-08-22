from __future__ import annotations

import tempfile
from pathlib import Path
import unittest
from unittest.mock import patch
import zipfile

from usable_xlsm.security import preflight_workbook


class FakeParser:
    has_vba = True
    has_xlm = False
    stomped = False
    analysis = []

    def __init__(self, _: str) -> None:
        pass

    def detect_vba_macros(self) -> bool:
        return self.has_vba

    def detect_xlm_macros(self) -> bool:
        return self.has_xlm

    def detect_vba_stomping(self) -> bool:
        return self.stomped

    def analyze_macros(self, **_: object) -> list[tuple[str, str, str]]:
        return list(self.analysis)

    def close(self) -> None:
        pass


class SecurityTests(unittest.TestCase):
    def make_book(self, root: Path, *, signed: bool = False) -> Path:
        workbook = root / "book.xlsm"
        with zipfile.ZipFile(workbook, "w") as package:
            package.writestr("[Content_Types].xml", "<Types/>")
            package.writestr("xl/vbaProject.bin", b"vba")
            if signed:
                package.writestr("xl/vbaProjectSignature.bin", b"signature")
        return workbook

    def setUp(self) -> None:
        FakeParser.has_vba = True
        FakeParser.has_xlm = False
        FakeParser.stomped = False
        FakeParser.analysis = []

    def test_execution_requires_explicit_trust(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            workbook = self.make_book(Path(raw))
            with patch("usable_xlsm.security.VBA_Parser", FakeParser):
                report = preflight_workbook(workbook, operation="execute")
                trusted = preflight_workbook(
                    workbook, operation="execute", trust_workbook=True
                )
        self.assertFalse(report.allowed)
        self.assertIn("source is not trusted", " ".join(report.blocking_reasons))
        self.assertTrue(trusted.allowed)

    def test_xlm_is_a_hard_blocker(self) -> None:
        FakeParser.has_xlm = True
        with tempfile.TemporaryDirectory() as raw:
            workbook = self.make_book(Path(raw))
            with patch("usable_xlsm.security.VBA_Parser", FakeParser):
                report = preflight_workbook(
                    workbook, operation="execute", trust_workbook=True
                )
        self.assertFalse(report.allowed)
        self.assertIn("Excel 4.0/XLM macros are not allowed", report.blocking_reasons)

    def test_signed_edit_requires_acknowledgement(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            workbook = self.make_book(Path(raw), signed=True)
            with patch("usable_xlsm.security.VBA_Parser", FakeParser):
                blocked = preflight_workbook(
                    workbook, operation="edit", trust_workbook=True
                )
                allowed = preflight_workbook(
                    workbook,
                    operation="edit",
                    trust_workbook=True,
                    allow_signature_removal=True,
                )
        self.assertFalse(blocked.allowed)
        self.assertTrue(allowed.allowed)

    def test_stomping_is_a_hard_blocker(self) -> None:
        FakeParser.stomped = True
        with tempfile.TemporaryDirectory() as raw:
            workbook = self.make_book(Path(raw))
            with patch("usable_xlsm.security.VBA_Parser", FakeParser):
                report = preflight_workbook(
                    workbook, operation="execute", trust_workbook=True
                )
        self.assertFalse(report.allowed)
        self.assertIn("VBA stomping was detected", report.blocking_reasons)


if __name__ == "__main__":
    unittest.main()
