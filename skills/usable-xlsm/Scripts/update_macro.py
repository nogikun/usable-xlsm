from pathlib import Path
import sys

import typer

from usable_xlsm import (
    DEFAULT_TIMEOUT,
    VbaModuleMismatchError,
    VbaSyntaxError,
    VbaUpdateError,
    WorkbookSecurityError,
    update_vba,
)


# Excel's dialogs and VBA errors are localised, so force UTF-8 on the streams:
# otherwise Windows encodes them in the console codepage and any tool reading
# this output gets mojibake instead of the error message.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")


app = typer.Typer(add_completion=False)


@app.command()
def update(
    input_path: Path = typer.Option(
        ...,
        "--input",
        "-i",
        exists=True,
        file_okay=False,
        dir_okay=True,
        help="Directory containing exported VBA files.",
    ),
    output_path: Path = typer.Option(
        ...,
        "--output",
        "-o",
        exists=True,
        file_okay=True,
        dir_okay=False,
        help="Existing XLSM file to update.",
    ),
    sync: bool = typer.Option(
        False,
        "--sync",
        help="Add modules that are missing and remove ones no longer present.",
    ),
    check: bool = typer.Option(
        True,
        "--check/--no-check",
        help="Run the syntax check before writing. Skipping risks a hung Excel.",
    ),
    backup: bool = typer.Option(
        True,
        "--backup/--no-backup",
        help="Copy the workbook first and restore it if the update fails.",
    ),
    trust_workbook: bool = typer.Option(
        False,
        "--trust-workbook",
        help="Explicitly attest that this exact workbook is trusted for this operation.",
    ),
    policy: Path | None = typer.Option(
        None,
        "--policy",
        exists=True,
        dir_okay=False,
        help="TOML policy containing approved roots or workbook hashes.",
    ),
    allow_signature_removal: bool = typer.Option(
        False,
        "--allow-signature-removal",
        help="Allow editing a signed workbook; an approved re-signing step is then required.",
    ),
    test: bool = typer.Option(
        True,
        "--test/--no-test",
        help="Run the staged workbook's VBA suite before promoting it.",
    ),
    require_tests: bool = typer.Option(
        True,
        "--require-tests/--allow-no-tests",
        help="Fail promotion when no Test_* procedures are discovered.",
    ),
    isolate_tests: bool = typer.Option(
        True,
        "--isolate-tests/--shared-test-copy",
        help="Give each test a fresh workbook copy to prevent state leakage.",
    ),
    audit_log: Path | None = typer.Option(
        None,
        "--audit-log",
        dir_okay=False,
        help="JSONL audit destination; defaults beside the workbook.",
    ),
    backup_keep: int = typer.Option(
        10,
        "--backup-keep",
        min=1,
        help="Number of newest managed backups to retain after success.",
    ),
    timeout: float = typer.Option(
        DEFAULT_TIMEOUT,
        "--timeout",
        "-t",
        help="Seconds before Excel is force-terminated.",
    ),
) -> None:
    try:
        report = update_vba(
            input_path,
            output_path,
            sync=sync,
            check=check,
            backup=backup,
            timeout=timeout,
            trust_workbook=trust_workbook,
            policy_path=policy,
            allow_signature_removal=allow_signature_removal,
            run_test_suite=test,
            require_tests=require_tests,
            isolate_tests=isolate_tests,
            audit_log=audit_log,
            backup_keep=backup_keep,
        )
    except VbaSyntaxError as exc:
        typer.echo(str(exc), err=True)
        typer.echo("\nNothing was written. Fix the errors above and retry.", err=True)
        raise typer.Exit(code=1) from exc
    except (
        VbaModuleMismatchError,
        VbaUpdateError,
        WorkbookSecurityError,
        ValueError,
        NotADirectoryError,
    ) as exc:
        typer.echo(f"Error: {exc}", err=True)
        raise typer.Exit(code=1) from exc

    typer.echo(f"Updated: {output_path}")
    for label in ("updated", "added", "removed"):
        names = report.get(label) or []
        if names:
            typer.echo(f"  {label}: {', '.join(names)}")
    if report.get("backup"):
        typer.echo(f"  backup: {report['backup']}")
    typer.echo(f"  sha256: {report['sha256_after']}")
    typer.echo(f"  tests: {report['tests']}")
    typer.echo(f"  audit: {report['audit_log']}")


if __name__ == "__main__":
    app(args=["--help"] if len(sys.argv) == 1 else None)
