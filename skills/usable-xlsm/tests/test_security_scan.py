from pathlib import Path
import json
import os
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import time

import oletools.olevba as olevba
from usable_xlsm._security_worker import _candidate_vba_strings
from usable_xlsm.security import preflight_workbook, WorkbookSecurityError
from usable_xlsm.core import update_vba


class SecurityScanTests(unittest.TestCase):
    def test_japanese_findings_ignore_parent_pipe_encoding(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'japanese.bas'
            source.write_text('Public Sub Auto_Open()\nx = Chr(65) & "日本語"\nEnd Sub\n', encoding='utf-8')
            with patch.dict(os.environ, {'PYTHONIOENCODING': 'cp932', 'PYTHONUTF8': '0'}):
                report = preflight_workbook(source, operation='inspect')
            self.assertNotIn('macro_scan_worker_failed', [item.code for item in report.findings])
            self.assertTrue(any('日本語' in item.summary for item in report.findings))

    def test_candidate_lines_preserve_all_scanner_findings(self):
        expressions = [
            'Public Function Value() As Long\nValue = 42\nEnd Function',
            'x = "hel" & "lo"',
            'x = Chr(104) & Chr$(116) & ChrB(116) & ChrW$(112)',
            'x = Chr(Asc("A") + Val("1"))',
            'x = StrReverse("exe.llehsrewop")',
            'x = Environ("TEMP")',
            'x = DecodeHex("687474703a2f2f6578616d706c652e636f6d")',
            'x = DecodeBase64("aHR0cDovL2V4YW1wbGUuY29t")',
            '\tx = ("http://" & _\n "example.com")',
            'x = "escaped ""quoted"" string" & " tail"',
            "' Chr(65) & Chr(66) in a comment",
            'Public Sub Auto_Open()\nShell "cmd.exe"\nEnd Sub',
            'x = CreateObject("WScript.Shell")',
            'x = cHrW ( &H41 ) & ChrB$(66)',
            'numeric = 1: value = (1 + 2) * 3',
        ]
        # Compare the complete upstream scan, not just decoded strings: this
        # catches changes to AutoExec/IOC/suspicious/hex/base64 findings too.
        for source in expressions + ['\n'.join(expressions)]:
            with self.subTest(source=source):
                expected = olevba.VBA_Scanner(source).scan(True, True)
                with patch('oletools.olevba.detect_vba_strings', _candidate_vba_strings):
                    actual = olevba.VBA_Scanner(source).scan(True, True)
                self.assertEqual(actual, expected)

    def test_unknown_olevba_version_uses_full_input(self):
        with patch('oletools.olevba.__version__', 'future-version'), patch(
            'usable_xlsm._security_worker._original_detect_strings', return_value=[]) as original:
            _candidate_vba_strings('value = 42\n')
        original.assert_called_once_with('value = 42\n')

    def test_timeout_kills_real_scanner_and_blocks_update_even_with_failure_policy(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            book = root / 'book.xlsm'
            book.write_bytes(b'disposable timeout fixture')
            before = book.read_bytes()
            policy = root / 'policy.toml'
            policy.write_text('[security]\nallow_scan_failures = true\n', encoding='utf-8')
            marker = root / 'should-not-exist'
            original_run = subprocess.run
            captured = []
            def stalled(command, **kwargs):
                captured.append(command[0])
                child = 'import time,pathlib; time.sleep(0.4); pathlib.Path(' + repr(str(marker)) + ').write_text("alive")'
                return original_run([command[0], '-c', child], **kwargs)
            with patch('usable_xlsm.security.subprocess.run', side_effect=stalled):
                report = preflight_workbook(book, operation='edit', trust_workbook=True,
                                            policy_path=policy, scan_timeout=0.15)
            self.assertFalse(report.allowed)
            self.assertIn('macro_scan_timeout', [item.code for item in report.findings])
            self.assertNotIn('no_vba', [item.code for item in report.findings])
            self.assertEqual(captured, [getattr(sys, '_base_executable', sys.executable)])
            time.sleep(0.5)
            self.assertFalse(marker.exists(), 'Scanner survived its timeout')
            with patch('usable_xlsm.core.preflight_workbook', return_value=report), patch(
                'usable_xlsm.core.run_job') as excel, patch('usable_xlsm.core.backup_workbook') as backup:
                with self.assertRaises(WorkbookSecurityError):
                    update_vba(root, book, partial=True, skip_unchanged=True, trust_workbook=True)
            excel.assert_not_called()
            backup.assert_not_called()
            self.assertEqual(book.read_bytes(), before)

    def test_crash_and_malformed_result_fail_closed(self):
        with tempfile.TemporaryDirectory() as directory:
            book = Path(directory) / 'book.xlsm'
            book.write_bytes(b'disposable failure fixture')
            invalid = [
                subprocess.CompletedProcess([], 1, stdout=b'', stderr=b'scanner failed'),
                subprocess.CompletedProcess([], 0, stdout=b'not json', stderr=b''),
                subprocess.CompletedProcess([], 0, stdout=json.dumps({
                    'has_vba': 'true', 'has_xlm': False, 'stomping': False,
                    'findings': [], 'blockers': []}).encode('utf-8'), stderr=b''),
                subprocess.CompletedProcess([], 0, stdout=json.dumps({
                    'has_vba': True, 'has_xlm': False, 'stomping': False,
                    'findings': [{'code': 'sample', 'severity': 'info',
                                  'summary': 'sample', 'detail': {'invalid': True}}],
                    'blockers': []}).encode('utf-8'), stderr=b''),
            ]
            for result in invalid:
                with self.subTest(result=result), patch('usable_xlsm.security.subprocess.run', return_value=result):
                    report = preflight_workbook(book, trust_workbook=True)
                    self.assertFalse(report.allowed)
                    self.assertIn('macro_scan_worker_failed', [item.code for item in report.findings])
            for timeout in (0, -1, float('inf'), float('nan')):
                with self.subTest(timeout=timeout), self.assertRaises(ValueError):
                    preflight_workbook(book, scan_timeout=timeout)

    def test_real_preflight_finishes_for_numeric_vba(self):
        # A text input keeps this regression focused on the actual macro scanner;
        # the unsupported workbook type must still prevent execution.
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'large.bas'
            source.write_text('\n'.join(
                f'Public Function Value{i}() As Long\nValue{i} = {i}\nEnd Function'
                for i in range(1000)), encoding='utf-8')
            command = [getattr(sys, '_base_executable', sys.executable), '-c',
                'import json,sys; sys.path=json.loads(sys.argv[1]); '
                'from usable_xlsm.security import preflight_workbook; '
                'print(json.dumps(preflight_workbook(sys.argv[2], operation="inspect").to_dict()))',
                json.dumps(sys.path), str(source)]
            try:
                result = subprocess.run(command, capture_output=True, text=True, timeout=5)
            except subprocess.TimeoutExpired:
                self.fail('Numeric VBA still spends over 5 seconds in security preflight')
            self.assertEqual(result.returncode, 0, result.stderr)
            report = json.loads(result.stdout)
            self.assertTrue(report['has_vba'])
            self.assertFalse(report['allowed'])
            self.assertTrue(any('unsupported workbook type' in reason for reason in report['blocking_reasons']))
            self.assertFalse(any('scan' in reason for reason in report['blocking_reasons']))


if __name__ == '__main__':
    unittest.main()
