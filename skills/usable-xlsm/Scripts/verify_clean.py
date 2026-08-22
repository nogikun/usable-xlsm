"""Check that the test harness left nothing behind.

Run this after a test session. The harness injects UsableXlsm* modules into the
workbook and duplicates it on disk; both are meant to disappear, and this reads
the real .xlsm and the real directory rather than taking that on trust.
"""

from pathlib import Path
import sys

import typer

from usable_xlsm import discard_working_copy, verify_clean


# Excel's dialogs and VBA errors are localised, so force UTF-8 on the streams:
# otherwise Windows encodes them in the console codepage and any tool reading
# this output gets mojibake instead of the error message.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")


app = typer.Typer(add_completion=False)


@app.command()
def verify(
    target: Path = typer.Option(
        ...,
        "--target",
        "-t",
        exists=True,
        help="Workbook to check, or a directory of workbooks.",
    ),
    fix: bool = typer.Option(
        False,
        "--fix",
        help="Delete any scratch copies found. Harness modules cannot be removed "
        "from here - rerun the tests, which cleans up as it closes.",
    ),
    include_backups: bool = typer.Option(
        False,
        "--include-backups",
        help="Also report *.bak.xlsm files. They are created deliberately, so "
        "they are not treated as leftovers by default.",
    ),
    skip_processes: bool = typer.Option(
        False,
        "--skip-processes",
        help="Do not report running Excel processes (they may be the user's own).",
    ),
) -> None:
    reports = verify_clean(target, check_processes=not skip_processes)

    # Scratch copies are a directory-level fact that every report repeats (one
    # workbook or ten, the same stray file shows up in each), so they are
    # tallied once here rather than through report.problems(), which would
    # double- or triple-count a single leftover file per workbook checked.
    workbook_problems = 0
    seen_scratch: set[str] = set()

    for report in reports:
        label = report.workbook.name if report.workbook else str(target)
        issues = [
            issue
            for issue in report.problems()
            if not issue.startswith("scratch copy left on disk:")
        ]
        if issues:
            workbook_problems += len(issues)
            typer.echo(f"NG {label}")
            for issue in issues:
                typer.echo(f"     {issue}")
        else:
            typer.echo(f"OK {label} - no harness modules")

        if include_backups and report.backups:
            typer.echo(f"   backups (kept on purpose): {', '.join(report.backups)}")

        seen_scratch.update(report.scratch_copies)

    for name in sorted(seen_scratch):
        typer.echo(f"NG scratch copy left on disk: {name}")

    scratch_problems = len(seen_scratch)

    if fix and seen_scratch:
        directory = target if target.is_dir() else target.parent
        for name in sorted(seen_scratch):
            if discard_working_copy(directory / name):
                typer.echo(f"   removed {name}")
                scratch_problems -= 1
            else:
                typer.echo(f"   could not remove {name} - Excel may still hold it")

    problems = workbook_problems + scratch_problems

    processes = reports[0].excel_processes if reports else []
    if processes:
        typer.echo(f"!  Excel still running: {sorted(processes)}")
        typer.echo("   Production jobs require a clean dedicated worker.")
        typer.echo("   Do not kill unidentified Excel processes; recycle the worker if needed.")

    if problems > 0:
        typer.echo(f"\n{problems} leftover(s) found.", err=True)
        raise typer.Exit(code=1)

    typer.echo("\nClean.")


if __name__ == "__main__":
    app(args=["--help"] if len(sys.argv) == 1 else None)
