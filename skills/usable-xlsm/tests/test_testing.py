from __future__ import annotations

from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from usable_xlsm.core import _run_test_discovery, run_tests
from usable_xlsm.runner import JobResult
from usable_xlsm.testing import Discovery, TestCase as VbaTestCase, discover, generate_runner


class TestingHarnessTests(unittest.TestCase):
    def test_teardown_runs_before_result_is_recorded(self) -> None:
        discovery = Discovery(
            cases=[VbaTestCase("Tests", "Test_One")],
            setups={},
            teardowns={"Tests": True},
            skipped=[],
        )
        source = generate_runner(discovery)
        teardown = source.index("Call Tests.TestTeardown")
        record = source.index('out = out & Record("Tests", "Test_One"')
        self.assertLess(teardown, record)
        self.assertIn('failurePhase = "TestTeardown"', source)
        self.assertIn("savedNumber = Err.Number", source)

    def test_no_tests_is_failure_when_required(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            workbook = Path(raw) / "book.xlsm"
            workbook.write_bytes(b"placeholder")
            with patch("usable_xlsm.core.extract_vba", return_value={}):
                run = run_tests(
                    workbook,
                    require_tests=True,
                    _skip_preflight=True,
                    _audit=False,
                )
        self.assertFalse(run.ok)
        self.assertEqual(run.error_code, "no_tests")

    def test_cleanup_failure_fails_the_run(self) -> None:
        discovery = discover(
            {"Tests.bas": "Public Sub Test_One()\nEnd Sub"}
        )
        with tempfile.TemporaryDirectory() as raw:
            workbook = Path(raw) / "book.xlsm"
            workbook.write_bytes(b"placeholder")
            with (
                patch(
                    "usable_xlsm.core.run_job",
                    return_value=JobResult(
                        ok=True,
                        data={"raw": "Tests\tTest_One\tpass\t0.001\t\n"},
                    ),
                ),
                patch("usable_xlsm.core.discard_working_copy", return_value=False),
            ):
                run = _run_test_discovery(
                    workbook, discovery, timeout=1, job_id="job"
                )
        self.assertFalse(run.ok)
        self.assertEqual(run.error_code, "cleanup_failed")


if __name__ == "__main__":
    unittest.main()
