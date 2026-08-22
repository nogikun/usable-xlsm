"""Installed command-line interface for usable-xlsm."""

from __future__ import annotations

import json
from pathlib import Path
import sys
from typing import Any

import typer

from .core import (
    check_syntax,
    extract_vba,
    restore_workbook,
    run_macro,
    run_tests,
    update_vba,
    verify_clean,
)
from .runner import DEFAULT_TIMEOUT, excel_pids
from .security import WorkbookSecurityError, preflight_workbook
from .support import create_support_bundle
from .syntax import VbaSyntaxError


app = typer.Typer(add_completion=False, no_args_is_help=True)


def _dump(value: Any) -> None:
    typer.echo(json.dumps(value, ensure_ascii=False, indent=2, default=str))


def _coerce(value: str) -> object:
    try:
        return json.loads(value)
    except json.JSONDecodeError:
        return value


@app.command("doctor")
def doctor_command() -> None:
    checks: dict[str, Any] = {"platform": sys.platform}
    settings: list[dict[str, Any]] = []
    if sys.platform == "win32":
        import winreg

        for version in ("16.0", "15.0", "14.0"):
            key_path = rf"Software\Microsoft\Office\{version}\Excel\Security"
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                    value, _ = winreg.QueryValueEx(key, "AccessVBOM")
                    settings.append({"office": version, "enabled": int(value) == 1})
            except OSError:
                continue
    checks["access_vbom"] = settings
    checks["excel_pids"] = sorted(excel_pids())
    try:
        import win32com.client  # noqa: F401
        checks["pywin32"] = True
    except ImportError:
        checks["pywin32"] = False
    try:
        import antlr4_vba  # noqa: F401
        checks["antlr4_vba"] = True
    except ImportError:
        checks["antlr4_vba"] = False
    checks["ok"] = bool(
        sys.platform == "win32"
        and any(item["enabled"] for item in settings)
        and checks["pywin32"]
        and checks["antlr4_vba"]
        and not checks["excel_pids"]
    )
    _dump(checks)
    if not checks["ok"]:
        raise typer.Exit(code=1)


@app.command()
def preflight(
    workbook: Path = typer.Option(..., "--workbook", "-w", exists=True, dir_okay=False),
    operation: str = typer.Option("execute", "--operation"),
    trust_workbook: bool = typer.Option(False, "--trust-workbook"),
    policy: Path | None = typer.Option(None, "--policy", exists=True, dir_okay=False),
    allow_signature_removal: bool = typer.Option(False, "--allow-signature-removal"),
) -> None:
    report = preflight_workbook(
        workbook,
        operation=operation,
        trust_workbook=trust_workbook,
        policy_path=policy,
        allow_signature_removal=allow_signature_removal,
    )
    _dump(report.to_dict())
    if not report.allowed:
        raise typer.Exit(code=2)


@app.command("extract")
def extract_command(
    workbook: Path = typer.Option(..., "--workbook", "-w", exists=True, dir_okay=False),
    output: Path = typer.Option(..., "--output", "-o", file_okay=False),
) -> None:
    output.mkdir(parents=True, exist_ok=True)
    modules = extract_vba(workbook)
    for name, source in modules.items():
        (output / Path(name).name).write_text(source, encoding="utf-8", newline="")
    _dump({"ok": True, "modules": sorted(Path(name).name for name in modules), "output": output})


@app.command("check")
def check_command(
    source: Path = typer.Option(..., "--source", "-s", exists=True, file_okay=False),
) -> None:
    issues = check_syntax(source)
    _dump(
        {
            "ok": not issues,
            "issues": [
                {"file": str(issue.path), "line": issue.line, "column": issue.column, "message": issue.message}
                for issue in issues
            ],
        }
    )
    if issues:
        raise typer.Exit(code=1)


@app.command("update")
def update_command(
    source: Path = typer.Option(..., "--source", "-s", exists=True, file_okay=False),
    workbook: Path = typer.Option(..., "--workbook", "-w", exists=True, dir_okay=False),
    trust_workbook: bool = typer.Option(False, "--trust-workbook"),
    policy: Path | None = typer.Option(None, "--policy", exists=True, dir_okay=False),
    sync: bool = typer.Option(False, "--sync"),
    allow_signature_removal: bool = typer.Option(False, "--allow-signature-removal"),
    test: bool = typer.Option(True, "--test/--no-test"),
    require_tests: bool = typer.Option(True, "--require-tests/--allow-no-tests"),
    isolate_tests: bool = typer.Option(True, "--isolate-tests/--shared-test-copy"),
    timeout: float = typer.Option(DEFAULT_TIMEOUT, "--timeout", "-t"),
    audit_log: Path | None = typer.Option(None, "--audit-log", dir_okay=False),
    backup_keep: int = typer.Option(10, "--backup-keep", min=1),
) -> None:
    try:
        report = update_vba(
            source,
            workbook,
            sync=sync,
            trust_workbook=trust_workbook,
            policy_path=policy,
            allow_signature_removal=allow_signature_removal,
            run_test_suite=test,
            require_tests=require_tests,
            isolate_tests=isolate_tests,
            timeout=timeout,
            audit_log=audit_log,
            backup_keep=backup_keep,
        )
    except (WorkbookSecurityError, VbaSyntaxError, ValueError, RuntimeError) as exc:
        _dump(
            {
                "ok": False,
                "error": str(exc),
                "error_code": getattr(exc, "error_code", type(exc).__name__),
            }
        )
        raise typer.Exit(code=2 if isinstance(exc, WorkbookSecurityError) else 1) from exc
    _dump({"ok": True, **report})


@app.command("test")
def test_command(
    workbook: Path = typer.Option(..., "--workbook", "-w", exists=True, dir_okay=False),
    trust_workbook: bool = typer.Option(False, "--trust-workbook"),
    policy: Path | None = typer.Option(None, "--policy", exists=True, dir_okay=False),
    pattern: str | None = typer.Option(None, "--filter", "-k"),
    require_tests: bool = typer.Option(True, "--require-tests/--allow-no-tests"),
    isolate: bool = typer.Option(True, "--isolate/--shared-copy"),
    timeout: float = typer.Option(DEFAULT_TIMEOUT, "--timeout", "-t"),
    audit_log: Path | None = typer.Option(None, "--audit-log", dir_okay=False),
) -> None:
    try:
        run = run_tests(
            workbook,
            pattern=pattern,
            trust_workbook=trust_workbook,
            policy_path=policy,
            require_tests=require_tests,
            isolate=isolate,
            timeout=timeout,
            audit_log=audit_log,
        )
    except WorkbookSecurityError as exc:
        _dump({"ok": False, "error": str(exc), "error_code": "security_preflight"})
        raise typer.Exit(code=2) from exc
    _dump(
        {
            "ok": run.ok,
            "job_id": run.job_id,
            "duration": run.duration,
            "error": run.error,
            "error_code": run.error_code,
            "leftover": run.leftover,
            "audit_log": run.audit_log,
            "office_versions": run.office_versions,
            "skipped": run.skipped,
            "results": [result.__dict__ for result in run.results],
        }
    )
    if not run.ok:
        raise typer.Exit(code=1)


@app.command("run")
def run_command(
    workbook: Path = typer.Option(..., "--workbook", "-w", exists=True, dir_okay=False),
    macro: str = typer.Option(..., "--macro", "-m"),
    arg: list[str] = typer.Option([], "--arg", "-a"),
    read_cell: list[str] = typer.Option([], "--read-cell", "-r"),
    trust_workbook: bool = typer.Option(False, "--trust-workbook"),
    policy: Path | None = typer.Option(None, "--policy", exists=True, dir_okay=False),
    on_copy: bool = typer.Option(True, "--on-copy/--in-place"),
    save: bool = typer.Option(False, "--save"),
    enable_events: bool = typer.Option(False, "--enable-events"),
    timeout: float = typer.Option(DEFAULT_TIMEOUT, "--timeout", "-t"),
    audit_log: Path | None = typer.Option(None, "--audit-log", dir_okay=False),
) -> None:
    try:
        result = run_macro(
            workbook,
            macro,
            args=[_coerce(value) for value in arg],
            read_cells=read_cell,
            on_copy=on_copy,
            allow_in_place=not on_copy,
            save=save,
            enable_events=enable_events,
            timeout=timeout,
            trust_workbook=trust_workbook,
            policy_path=policy,
            audit_log=audit_log,
        )
    except WorkbookSecurityError as exc:
        _dump({"ok": False, "error": str(exc), "error_code": "security_preflight"})
        raise typer.Exit(code=2) from exc
    _dump(
        {
            "ok": result.ok,
            "job_id": result.job_id,
            "duration": result.duration,
            "error": result.error,
            "error_code": result.error_code,
            "data": result.data if result.ok else None,
            "office_version": result.office_version,
        }
    )
    if not result.ok:
        raise typer.Exit(code=1)


@app.command("restore")
def restore_command(
    workbook: Path = typer.Option(..., "--workbook", "-w", dir_okay=False),
    backup: Path | None = typer.Option(None, "--backup", "-b", exists=True, dir_okay=False),
    audit_log: Path | None = typer.Option(None, "--audit-log", dir_okay=False),
) -> None:
    _dump(restore_workbook(workbook, backup_path=backup, audit_log=audit_log))


@app.command("verify")
def verify_command(
    target: Path = typer.Option(..., "--target", "-t", exists=True),
) -> None:
    reports = verify_clean(target)
    problems = [problem for report in reports for problem in report.problems()]
    _dump(
        {
            "ok": not problems,
            "problems": problems,
            "excel_pids": sorted(excel_pids()),
        }
    )
    if problems:
        raise typer.Exit(code=1)


@app.command("support-bundle")
def support_bundle_command(
    workbook: Path = typer.Option(..., "--workbook", "-w", exists=True, dir_okay=False),
    output: Path = typer.Option(..., "--output", "-o", dir_okay=False),
    audit_log: Path | None = typer.Option(None, "--audit-log", exists=True, dir_okay=False),
    max_events: int = typer.Option(100, "--max-events", min=0),
) -> None:
    path = create_support_bundle(
        workbook,
        output,
        audit_log=audit_log,
        max_events=max_events,
    )
    _dump({"ok": True, "support_bundle": str(path), "contains_workbook_data": False})


def main() -> None:
    app()


if __name__ == "__main__":
    main()
