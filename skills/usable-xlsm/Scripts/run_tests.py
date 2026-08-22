"""Discover and run the VBA tests inside a workbook.

A test is a Public Sub named Test_Something, taking no arguments, in a standard
module. Tests always run against a temporary copy of the workbook.
"""

import json
from pathlib import Path
import sys

import typer

from usable_xlsm import DEFAULT_TIMEOUT, WorkbookSecurityError, run_tests, verify_clean


# Excel's dialogs and VBA errors are localised, so force UTF-8 on the streams:
# otherwise Windows encodes them in the console codepage and any tool reading
# this output gets mojibake instead of the error message.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")


app = typer.Typer(add_completion=False)

_MARKS = {"pass": "ok", "fail": "FAIL", "error": "ERROR"}


@app.command()
def test(
    workbook: Path = typer.Option(
        ...,
        "--workbook",
        "-w",
        exists=True,
        dir_okay=False,
        help="Macro-enabled workbook containing the tests.",
    ),
    pattern: str = typer.Option(
        None,
        "--filter",
        "-k",
        help="Only run tests whose Module.Name matches this regular expression.",
    ),
    as_json: bool = typer.Option(False, "--json", help="Emit machine-readable results."),
    trust_workbook: bool = typer.Option(
        False,
        "--trust-workbook",
        help="Explicitly attest that this exact workbook is trusted for execution.",
    ),
    policy: Path | None = typer.Option(
        None,
        "--policy",
        exists=True,
        dir_okay=False,
        help="TOML policy containing approved roots or workbook hashes.",
    ),
    require_tests: bool = typer.Option(
        True,
        "--require-tests/--allow-no-tests",
        help="Exit non-zero when no Test_* procedures are discovered.",
    ),
    isolate: bool = typer.Option(
        True,
        "--isolate/--shared-copy",
        help="Run each test on a fresh workbook copy.",
    ),
    audit_log: Path | None = typer.Option(
        None,
        "--audit-log",
        dir_okay=False,
        help="JSONL audit destination; defaults beside the workbook.",
    ),
    timeout: float = typer.Option(
        DEFAULT_TIMEOUT,
        "--timeout",
        "-t",
        help="Seconds before Excel is force-terminated.",
    ),
) -> None:
    try:
        run = run_tests(
            workbook,
            pattern=pattern,
            timeout=timeout,
            trust_workbook=trust_workbook,
            policy_path=policy,
            require_tests=require_tests,
            isolate=isolate,
            audit_log=audit_log,
        )
    except WorkbookSecurityError as exc:
        typer.echo(str(exc), err=True)
        raise typer.Exit(code=2) from exc

    if as_json:
        typer.echo(
            json.dumps(
                {
                    "ok": run.ok,
                    "duration": round(run.duration, 3),
                    "error": run.error,
                    "dialog": run.dialog,
                    "timed_out": run.timed_out,
                    "leftover": run.leftover,
                    "error_code": run.error_code,
                    "job_id": run.job_id,
                    "audit_log": run.audit_log,
                    "skipped": run.skipped,
                    "results": [
                        {
                            "module": result.module,
                            "name": result.name,
                            "status": result.status,
                            "duration": result.duration,
                            "message": result.message,
                        }
                        for result in run.results
                    ],
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        raise typer.Exit(code=0 if run.ok else 1)

    for note in run.skipped:
        typer.echo(f"skipped: {note}", err=True)

    if run.leftover:
        typer.echo(
            f"warning: could not delete the scratch copy {run.leftover}\n"
            "         Excel may still be holding it; doctor.py --sweep removes it.",
            err=True,
        )

    if run.error:
        typer.echo(run.error, err=True)
        raise typer.Exit(code=1)

    if not run.results:
        typer.echo("No tests found. Name them Test_Something in a .bas module.")
        if require_tests:
            raise typer.Exit(code=1)
        return

    width = max(len(result.qualified) for result in run.results)
    for result in run.results:
        mark = _MARKS.get(result.status, result.status)
        line = f"{result.qualified.ljust(width)}  {mark}"
        if result.message:
            line += f"  {result.message}"
        typer.echo(line)

    cleanup_issues: list[str] = []
    for report in verify_clean(workbook, check_processes=False):
        for issue in report.problems():
            cleanup_issues.append(issue)
            typer.echo(f"error: {issue}", err=True)

    failed = run.failed
    typer.echo(
        f"\n{len(run.results)} test(s), {len(run.passed)} passed, "
        f"{len(failed)} failed  ({run.duration:.1f}s)"
    )
    if failed or cleanup_issues or not run.ok:
        raise typer.Exit(code=1)


if __name__ == "__main__":
    app(args=["--help"] if len(sys.argv) == 1 else None)
