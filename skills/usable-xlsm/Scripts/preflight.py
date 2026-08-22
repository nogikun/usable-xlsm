"""Statically inspect and authorize a macro-enabled workbook without Excel."""

import json
from pathlib import Path
import sys

import typer

from usable_xlsm import preflight_workbook


for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")


app = typer.Typer(add_completion=False)


@app.command()
def preflight(
    workbook: Path = typer.Option(
        ...,
        "--workbook",
        "-w",
        exists=True,
        dir_okay=False,
        help="Macro-enabled workbook to inspect without launching Excel.",
    ),
    operation: str = typer.Option(
        "execute",
        "--operation",
        help="Planned operation: inspect, edit, or execute.",
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
        help="TOML policy containing approved roots or exact hashes.",
    ),
    allow_signature_removal: bool = typer.Option(
        False,
        "--allow-signature-removal",
        help="Acknowledge that an edit invalidates embedded signatures.",
    ),
    as_json: bool = typer.Option(True, "--json/--text", help="Output JSON or concise text."),
) -> None:
    report = preflight_workbook(
        workbook,
        operation=operation,
        trust_workbook=trust_workbook,
        policy_path=policy,
        allow_signature_removal=allow_signature_removal,
    )
    if as_json:
        typer.echo(json.dumps(report.to_dict(), ensure_ascii=False, indent=2))
    else:
        typer.echo(f"{'ALLOW' if report.allowed else 'BLOCK'} {workbook.name} ({operation})")
        typer.echo(f"sha256: {report.sha256}")
        for finding in report.findings:
            typer.echo(f"{finding.severity}: {finding.code}: {finding.summary}")
        for reason in report.blocking_reasons:
            typer.echo(f"blocked: {reason}")
    if not report.allowed:
        raise typer.Exit(code=2)


if __name__ == "__main__":
    app(args=["--help"] if len(sys.argv) == 1 else None)
