"""Run a macro in an XLSM and report its result as JSON.

Every Excel call goes through the watchdog, so a compile error surfaces as a
reported failure with the dialog text rather than an invisible hung process.
"""

import json
from pathlib import Path
import sys

import typer

from usable_xlsm import DEFAULT_TIMEOUT, WorkbookSecurityError, run_macro


# Excel's dialogs and VBA errors are localised, so force UTF-8 on the streams:
# otherwise Windows encodes them in the console codepage and any tool reading
# this output gets mojibake instead of the error message.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")


app = typer.Typer(add_completion=False)


def _coerce(value: str) -> object:
    """Read an argument as JSON so numbers and booleans survive the CLI."""
    try:
        return json.loads(value)
    except json.JSONDecodeError:
        return value


@app.command()
def run(
    workbook: Path = typer.Option(
        ...,
        "--workbook",
        "-w",
        exists=True,
        dir_okay=False,
        help="XLSM file containing the macro.",
    ),
    macro: str = typer.Option(
        ...,
        "--macro",
        "-m",
        help="Macro to call, e.g. Module1.MyMacro.",
    ),
    arg: list[str] = typer.Option(
        [],
        "--arg",
        "-a",
        help="Argument to pass; repeat for several. Parsed as JSON when possible.",
    ),
    read_cell: list[str] = typer.Option(
        [],
        "--read-cell",
        "-r",
        help="Cell to read after the run, e.g. A1 or Sheet1!B2. Repeatable.",
    ),
    on_copy: bool = typer.Option(
        True,
        "--on-copy/--in-place",
        help="Use a disposable copy (production default); --in-place is an explicit exception.",
    ),
    save: bool = typer.Option(False, "--save", help="Save the workbook after running."),
    enable_events: bool = typer.Option(
        False,
        "--enable-events",
        help="Allow Worksheet_Change and similar event handlers to fire.",
    ),
    timeout: float = typer.Option(
        DEFAULT_TIMEOUT,
        "--timeout",
        "-t",
        help="Seconds before Excel is force-terminated.",
    ),
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
    audit_log: Path | None = typer.Option(
        None,
        "--audit-log",
        dir_okay=False,
        help="JSONL audit destination; defaults beside the workbook.",
    ),
) -> None:
    try:
        result = run_macro(
            workbook,
            macro,
            args=[_coerce(value) for value in arg],
            read_cells=list(read_cell),
            save=save,
            on_copy=on_copy,
            allow_in_place=not on_copy,
            enable_events=enable_events,
            timeout=timeout,
            trust_workbook=trust_workbook,
            policy_path=policy,
            audit_log=audit_log,
        )
    except WorkbookSecurityError as exc:
        typer.echo(str(exc), err=True)
        raise typer.Exit(code=2) from exc

    payload = {
        "ok": result.ok,
        "stage": "run",
        "macro": macro,
        "duration": round(result.duration, 3),
        "job_id": result.job_id,
        "error_code": result.error_code,
    }
    if result.ok:
        payload.update(result.data or {})
    else:
        payload.update(
            {"error": result.error, "dialog": result.dialog, "timed_out": result.timed_out}
        )

    typer.echo(json.dumps(payload, ensure_ascii=False, indent=2, default=str))
    if not result.ok:
        raise typer.Exit(code=1)


if __name__ == "__main__":
    app(args=["--help"] if len(sys.argv) == 1 else None)
