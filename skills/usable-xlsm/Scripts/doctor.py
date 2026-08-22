"""Check the environment and clean up Excel processes left behind by a hang."""

from pathlib import Path
import json
import platform
import sys
import winreg

import typer

from usable_xlsm import excel_pids, mark_of_the_web, sweep_working_copies


# Excel's dialogs and VBA errors are localised, so force UTF-8 on the streams:
# otherwise Windows encodes them in the console codepage and any tool reading
# this output gets mojibake instead of the error message.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")


app = typer.Typer(add_completion=False)


def _access_vbom() -> list[tuple[str, int]]:
    """Read the 'Trust access to the VBA project object model' flag per Office version.

    Without this, every write to a VBA project fails, so it is the first thing
    worth checking when an update errors out.
    """
    found: list[tuple[str, int]] = []
    for version in ("16.0", "15.0", "14.0"):
        path = rf"Software\Microsoft\Office\{version}\Excel\Security"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
                value, _ = winreg.QueryValueEx(key, "AccessVBOM")
                found.append((version, int(value)))
        except OSError:
            continue
    return found


@app.command()
def doctor(
    workbook: Path = typer.Option(
        None,
        "--workbook",
        "-w",
        exists=True,
        dir_okay=False,
        help="Also check this workbook for a Mark of the Web (blocked macros in Excel).",
    ),
    sweep: Path = typer.Option(
        None,
        "--sweep",
        exists=True,
        file_okay=False,
        help="Delete managed tests/testrun/staging scratch copies in this directory.",
    ),
    as_json: bool = typer.Option(False, "--json", help="Emit machine-readable diagnostics."),
) -> None:
    problems = 0
    checks: dict[str, object] = {"platform": platform.platform()}
    def emit(message: str) -> None:
        if not as_json:
            typer.echo(message)

    settings = _access_vbom()
    checks["access_vbom"] = [{"office": version, "enabled": value == 1} for version, value in settings]
    if not settings:
        emit("? AccessVBOM: no Excel security key found. Is Excel installed?")
        problems += 1
    elif any(value == 1 for _, value in settings):
        enabled = ", ".join(v for v, value in settings if value == 1)
        emit(f"OK AccessVBOM enabled for Office {enabled}")
    else:
        emit("NG AccessVBOM is disabled - VBA writes will fail.")
        emit("   Excel > File > Options > Trust Center > Trust Center Settings")
        emit("   > Macro Settings > Trust access to the VBA project object model")
        problems += 1

    try:
        import win32com.client  # noqa: F401

        checks["pywin32"] = True
        emit("OK pywin32 available")
    except ImportError:
        checks["pywin32"] = False
        emit("NG pywin32 missing - install project dependencies.")
        problems += 1

    try:
        from antlr4_vba.vbaParser import vbaParser  # noqa: F401

        checks["antlr4_vba"] = True
        emit("OK antlr4-vba available (syntax checking enabled)")
    except ImportError:
        checks["antlr4_vba"] = False
        emit("NG antlr4-vba missing - syntax checking unavailable.")
        problems += 1

    if workbook is not None:
        mark = mark_of_the_web(workbook)
        checks["workbook"] = {"name": workbook.name, "motw": mark is not None}
        if mark is None:
            emit(f"OK {workbook.name} has no Mark of the Web")
        else:
            zone = next(
                (line.split("=", 1)[1] for line in mark.splitlines() if line.startswith("ZoneId=")),
                "?",
            )
            emit(f"!  {workbook.name} carries a Mark of the Web (ZoneId={zone})")
            emit("   Excel will block its macros in the UI: 'a potentially dangerous")
            emit("   macro has been blocked'. Production commands also block this file")
            emit("   during preflight. Do not clear the mark automatically; move an")
            emit("   independently reviewed copy into an approved trust root instead.")
            problems += 1

    if sweep is not None:
        removed = sweep_working_copies(sweep)
        checks["sweep"] = [path.name for path in removed]
        if removed:
            for path in removed:
                emit(f"   removed {path.name}")
            emit(f"OK swept {len(removed)} scratch cop(ies) from {sweep}")
        else:
            emit(f"OK no scratch copies found in {sweep}")

    running = excel_pids()
    checks["excel_pids"] = sorted(running)
    if running:
        emit(f"!  {len(running)} Excel process(es) running: {sorted(running)}")
        emit("   Production jobs refuse to start until the dedicated worker is clean.")
        emit("   This command never kills unidentified Excel processes.")
        problems += 1
    else:
        emit("OK no Excel processes running")

    checks["ok"] = problems == 0
    if as_json:
        typer.echo(json.dumps(checks, ensure_ascii=False, indent=2))
    if problems:
        raise typer.Exit(code=1)


if __name__ == "__main__":
    # Unlike the other scripts, every option here is optional, so a bare
    # invocation should actually run the checks rather than print usage.
    app()
