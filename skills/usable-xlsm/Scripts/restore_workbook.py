"""Restore a managed usable-xlsm backup atomically."""

import json
from pathlib import Path
import sys

import typer

from usable_xlsm import restore_workbook


app = typer.Typer(add_completion=False)


@app.command()
def restore(
    workbook: Path = typer.Option(..., "--workbook", "-w", dir_okay=False),
    backup: Path | None = typer.Option(
        None,
        "--backup",
        "-b",
        exists=True,
        dir_okay=False,
        help="Backup to restore; defaults to the newest managed backup.",
    ),
    audit_log: Path | None = typer.Option(None, "--audit-log", dir_okay=False),
) -> None:
    report = restore_workbook(workbook, backup_path=backup, audit_log=audit_log)
    typer.echo(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    app(args=["--help"] if len(sys.argv) == 1 else None)
