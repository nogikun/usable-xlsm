"""Excel COM worker, always executed in a throwaway subprocess.

Everything here is deliberately isolated from the parent process. A VBA compile
error makes Excel raise a modal dialog that no COM call can dismiss, so the
worker becomes permanently blocked. The parent (``runner.py``) treats that as a
timeout and kills this process plus the Excel instance it spawned - which is
only safe because both are disposable.

Protocol: a JSON job on stdin, a single JSON result on the last line of stdout.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
import sys
import traceback
from typing import Any


def _excel_pid(excel: Any) -> int | None:
    """Resolve the PID of the Excel instance we just started.

    The parent needs this to kill the right process when a dialog blocks us;
    without it, a watchdog could only guess and might kill the user's own Excel.
    """
    try:
        import win32process

        hwnd = int(excel.Hwnd)
    except Exception:
        return None
    try:
        return int(win32process.GetWindowThreadProcessId(hwnd)[1])
    except Exception:
        return None


def _code_module_body(source: str) -> str:
    """Strip Attribute lines; Excel regenerates them from the component itself."""
    lines = source.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    body = [line for line in lines if not line.startswith("Attribute ")]
    return "\r\n".join(body).strip("\r\n")


def _set_module_code(component: Any, source: str) -> None:
    code_module = component.CodeModule
    if code_module.CountOfLines:
        code_module.DeleteLines(1, code_module.CountOfLines)
    body = _code_module_body(source)
    if body:
        code_module.AddFromString(body)


# vbext_ComponentType values; .frm is intentionally unsupported because a form
# also carries a binary designer that plain text cannot reconstruct.
_COMPONENT_TYPES = {".bas": 1, ".cls": 2}


def _apply_modules(project: Any, modules: dict[str, str], sync: bool) -> dict[str, list[str]]:
    components = project.VBComponents
    existing = {component.Name: component for component in components}
    report: dict[str, list[str]] = {"updated": [], "added": [], "removed": []}

    wanted: dict[str, str] = {}
    for filename, source in modules.items():
        wanted[Path(filename).stem] = filename

    for filename, source in modules.items():
        name = Path(filename).stem
        suffix = Path(filename).suffix.lower()
        if suffix == ".frm":
            raise ValueError(
                f"UserForm source cannot be safely round-tripped: {filename}. "
                "The paired binary .frx designer would be lost."
            )
        if name in existing:
            if int(existing[name].Type) == 3:
                raise ValueError(
                    f"Refusing to modify UserForm {name}; plain-text updates cannot "
                    "preserve its binary designer."
                )
            _set_module_code(existing[name], source)
            report["updated"].append(name)
        elif sync:
            kind = _COMPONENT_TYPES.get(suffix)
            if kind is None:
                raise ValueError(f"Cannot create a component for {filename}")
            component = components.Add(kind)
            component.Name = name
            _set_module_code(component, source)
            report["added"].append(name)
        else:
            raise ValueError(f"Module {name} does not exist in the workbook")

    if sync:
        for name, component in existing.items():
            # Document modules (sheets, ThisWorkbook) belong to the workbook and
            # cannot be removed even when absent from the source directory.
            if name not in wanted and int(component.Type) in _COMPONENT_TYPES.values():
                components.Remove(component)
                report["removed"].append(name)

    return report


def _collection_item(collection: Any, name: str) -> Any | None:
    try:
        return collection.Item(name)
    except Exception:
        try:
            return collection(name)
        except Exception:
            return None


def _worksheet(book: Any, name: str) -> Any | None:
    worksheets = getattr(book, "Worksheets")
    try:
        return worksheets(name) if callable(worksheets) else worksheets.Item(name)
    except Exception:
        try:
            return worksheets.Item(name)
        except Exception:
            return None


def _apply_worksheets(book: Any, specs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    worksheets = getattr(book, "Worksheets")
    visibility = {"visible": -1, "hidden": 0, "veryHidden": 2}
    applied: list[dict[str, Any]] = []
    for spec in specs:
        sheet = _worksheet(book, spec["name"])
        created = False
        if sheet is None:
            if not spec["create"]:
                raise ValueError(f"Worksheet {spec['name']} does not exist; set create=true to add it")
            try:
                count = int(worksheets.Count)
                after = worksheets.Item(count)
                sheet = worksheets.Add(After=after)
            except Exception as exc:
                raise ValueError(f"Could not create worksheet {spec['name']}: {exc}") from exc
            sheet.Name = spec["name"]
            created = True
        sheet.Visible = visibility[spec["visible"]]
        applied.append({"name": str(sheet.Name), "visible": spec["visible"], "created": created})
    return applied


def _range_is_nonempty(cell: Any) -> bool:
    try:
        formula = str(getattr(cell, "Formula") or "").strip()
        if formula:
            return True
    except Exception:
        pass
    try:
        value = getattr(cell, "Value")
    except Exception:
        value = None
    return value not in (None, "")


def _apply_cells(book: Any, specs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    applied: list[dict[str, Any]] = []
    for spec in specs:
        sheet = _worksheet(book, spec["sheet"])
        if sheet is None:
            raise ValueError(f"Worksheet {spec['sheet']} does not exist for cell {spec['address']}")
        cell = sheet.Range(spec["address"])
        if not spec["overwrite"] and _range_is_nonempty(cell):
            raise ValueError(
                f"Cell {spec['sheet']}!{spec['address']} is not empty; set overwrite=true to replace it"
            )
        if spec["formula"] is not None:
            cell.Formula = spec["formula"]
        else:
            cell.Value = spec["value"]
        if spec["number_format"] is not None:
            cell.NumberFormat = spec["number_format"]
        applied.append(
            {
                "sheet": spec["sheet"],
                "address": spec["address"],
                "formula": spec["formula"],
                "value": spec["value"] if spec["formula"] is None else None,
                "number_format": spec["number_format"],
            }
        )
    return applied


def _verify_worksheets(book: Any, specs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    visibility = {"visible": -1, "hidden": 0, "veryHidden": 2}
    verified: list[dict[str, Any]] = []
    for spec in specs:
        sheet = _worksheet(book, spec["name"])
        if sheet is None:
            raise ValueError(f"Worksheet {spec['name']} was not found after save")
        actual = int(sheet.Visible)
        if actual != visibility[spec["visible"]]:
            raise ValueError(
                f"Worksheet {spec['name']} visibility mismatch: "
                f"expected={spec['visible']!r}, actual={actual}"
            )
        verified.append({"name": spec["name"], "verified": True})
    return verified


def _verify_cells(book: Any, specs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    verified: list[dict[str, Any]] = []
    for spec in specs:
        sheet = _worksheet(book, spec["sheet"])
        if sheet is None:
            raise ValueError(f"Worksheet {spec['sheet']} was not found after save")
        cell = sheet.Range(spec["address"])
        if spec["formula"] is not None:
            actual_formula = str(getattr(cell, "Formula") or "")
            if actual_formula != spec["formula"]:
                raise ValueError(
                    f"Cell {spec['sheet']}!{spec['address']} formula mismatch: "
                    f"expected={spec['formula']!r}, actual={actual_formula!r}"
                )
        else:
            actual = getattr(cell, "Value")
            expected = spec["value"]
            if isinstance(expected, bool):
                matches = actual is expected
            elif isinstance(expected, (int, float)) and not isinstance(expected, bool):
                try:
                    matches = abs(float(actual) - float(expected)) < 1e-9
                except (TypeError, ValueError):
                    matches = False
            else:
                matches = actual == expected
            if not matches:
                raise ValueError(
                    f"Cell {spec['sheet']}!{spec['address']} value mismatch: "
                    f"expected={expected!r}, actual={actual!r}"
                )
        if spec["number_format"] is not None:
            actual_format = str(getattr(cell, "NumberFormat") or "")
            if actual_format != spec["number_format"]:
                raise ValueError(
                    f"Cell {spec['sheet']}!{spec['address']} number format mismatch: "
                    f"expected={spec['number_format']!r}, actual={actual_format!r}"
                )
        verified.append({"sheet": spec["sheet"], "address": spec["address"], "verified": True})
    return verified


def _worksheet_buttons(sheet: Any) -> Any:
    buttons = getattr(sheet, "Buttons")
    return buttons() if callable(buttons) else buttons


def _apply_buttons(book: Any, specs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    applied: list[dict[str, Any]] = []
    for spec in specs:
        sheet = _worksheet(book, spec["sheet"])
        if sheet is None:
            raise ValueError(f"Worksheet {spec['sheet']} does not exist for button {spec['name']}")
        buttons = _worksheet_buttons(sheet)
        existing = _collection_item(buttons, spec["name"])
        replaced = False
        if existing is not None:
            if not spec["replace"]:
                raise ValueError(
                    f"Button {spec['name']} already exists on sheet {spec['sheet']}; "
                    "set replace=true to replace it."
                )
            existing.Delete()
            replaced = True

        button = buttons.Add(
            spec["left"],
            spec["top"],
            spec["width"],
            spec["height"],
        )
        button.Name = spec["name"]
        button.Caption = spec["caption"]
        button.OnAction = spec["macro"]
        applied.append(
            {
                "sheet": spec["sheet"],
                "name": str(button.Name),
                "caption": spec["caption"],
                "macro": spec["macro"],
                "left": spec["left"],
                "top": spec["top"],
                "width": spec["width"],
                "height": spec["height"],
                "replaced": replaced,
            }
        )
    return applied


def _verify_buttons(book: Any, specs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    verified: list[dict[str, Any]] = []
    for spec in specs:
        sheet = _worksheet(book, spec["sheet"])
        if sheet is None:
            raise ValueError(f"Worksheet {spec['sheet']} does not exist for button {spec['name']}")
        button = _collection_item(_worksheet_buttons(sheet), spec["name"])
        if button is None:
            raise ValueError(f"Button {spec['name']} was not found after save")
        expected_fields = {"Name": "name", "Caption": "caption", "OnAction": "macro"}
        for field, spec_field in expected_fields.items():
            expected = spec[spec_field]
            if str(getattr(button, field)) != str(expected):
                raise ValueError(
                    f"Button {spec['name']} {field} mismatch after save: "
                    f"expected={expected!r}, actual={getattr(button, field)!r}"
                )
        for field in ("Left", "Top", "Width", "Height"):
            expected = float(spec[field.lower()])
            actual = float(getattr(button, field))
            # Excel stores shape geometry in point/column units and rounds the
            # requested values on save.  A one-point tolerance validates the
            # persisted placement without rejecting normal COM round-off.
            if abs(actual - expected) > 1.0:
                raise ValueError(
                    f"Button {spec['name']} {field} mismatch after save: "
                    f"expected={expected}, actual={actual}"
                )
        verified.append({"sheet": spec["sheet"], "name": spec["name"], "verified": True})
    return verified


def _job_update(excel: Any, book: Any, job: dict[str, Any]) -> dict[str, Any]:
    project = book.VBProject
    report: dict[str, Any] = {"updated": [], "added": [], "removed": []}
    if job.get("modules"):
        report.update(_apply_modules(project, job["modules"], bool(job.get("sync"))))
    if job.get("worksheets"):
        report["worksheets"] = _apply_worksheets(book, job["worksheets"])
    if job.get("cells"):
        report["cells"] = _apply_cells(book, job["cells"])
    if job.get("buttons"):
        report["buttons"] = _apply_buttons(book, job["buttons"])

    # Run finalizers while the workbook is still the staging copy. This covers
    # trusted workbook objects that are outside the declarative script contract
    # without opening the user's original workbook for a second in-place write.
    post_macro = str(job.get("post_macro") or "").strip()
    if post_macro:
        excel.Run(f"'{book.Name}'!{post_macro}", *(job.get("post_macro_args") or []))
        report["post_macro"] = post_macro

    book.Save()
    if job.get("worksheets"):
        report["worksheets_verified"] = _verify_worksheets(book, job["worksheets"])
    if job.get("cells"):
        report["cells_verified"] = _verify_cells(book, job["cells"])
    if job.get("buttons"):
        report["buttons_verified"] = _verify_buttons(book, job["buttons"])
    report["components"] = [
        {"name": component.Name, "type": int(component.Type)}
        for component in project.VBComponents
    ]
    return report


def _job_run(excel: Any, book: Any, job: dict[str, Any]) -> dict[str, Any]:
    macro = job["macro"]
    args = job.get("args") or []
    result = excel.Run(f"'{book.Name}'!{macro}", *args)

    cells: dict[str, Any] = {}
    for reference in job.get("read_cells") or []:
        sheet_name, _, address = reference.rpartition("!")
        sheet = book.Worksheets(sheet_name) if sheet_name else book.Worksheets(1)
        cells[reference] = sheet.Range(address).Value

    if job.get("save"):
        book.Save()

    return {"result": result, "cells": cells}


def _job_run_tests(excel: Any, book: Any, job: dict[str, Any]) -> dict[str, Any]:
    """Inject the harness modules, run the generated entry point, remove them.

    This always operates on a temporary copy of the workbook (see
    core.run_tests), so the injected scaffolding can never reach the file the
    user is developing.
    """
    project = book.VBProject
    components = project.VBComponents
    injected: list[str] = []

    try:
        for name, code in job["modules"].items():
            # The runner is regenerated per run, so a stale one must go. A module
            # the user vendored themselves is left alone - see job["replace"].
            if name in job.get("replace", []):
                try:
                    components.Remove(components.Item(name))
                except Exception:
                    pass
            elif any(component.Name == name for component in components):
                continue

            component = components.Add(1)
            component.Name = name
            component.CodeModule.AddFromString(_code_module_body(code))
            injected.append(name)

        raw = excel.Run(f"'{book.Name}'!{job['entry']}")
        return {"raw": raw or "", "injected": injected}
    finally:
        # The copy is discarded and never saved, so this is belt-and-braces -
        # but scaffolding that removes itself is the only version that stays
        # safe if this is ever pointed at a workbook someone keeps.
        for name in injected:
            try:
                components.Remove(components.Item(name))
            except Exception:
                pass


def _job_inspect(book: Any, job: dict[str, Any]) -> dict[str, Any]:
    project = book.VBProject
    return {
        "project": project.Name,
        "components": [
            {"name": component.Name, "type": int(component.Type)}
            for component in project.VBComponents
        ],
        "sheets": [sheet.Name for sheet in book.Worksheets],
    }


_JOBS = {
    "update": _job_update,
    "run": _job_run,
    "run_tests": _job_run_tests,
    "inspect": lambda excel, book, job: _job_inspect(book, job),
}


def main() -> int:
    job = json.load(sys.stdin)
    pid_path = job.get("pid_file")
    job_id = str(job.get("job_id") or "")
    workbook = str(Path(job["workbook"]).resolve())

    if sys.platform != "win32":
        print(
            json.dumps(
                {
                    "ok": False,
                    "job_id": job_id,
                    "error_code": "excel_backend_unavailable",
                    "error": (
                        "Excel COM automation is only supported on Windows; "
                        "use the macOS manual backend for manifest/source validation."
                    ),
                }
            )
        )
        return 1

    import win32com.client

    excel = None
    book = None
    excel_pid = None
    previous_security = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = bool(job.get("enable_events", False))
        excel.AskToUpdateLinks = False
        excel.UserControl = False

        # Publish the PID before touching the VBA project, because everything
        # after this point is capable of hanging on a modal dialog.
        if pid_path:
            excel_pid = _excel_pid(excel)
            if excel_pid is not None:
                Path(pid_path).write_text(
                    json.dumps(
                        {
                            "job_id": job_id,
                            "worker_pid": os.getpid(),
                            "excel_pid": excel_pid,
                        }
                    ),
                    encoding="utf-8",
                )

        # Security is explicit for every open. Editing/inspection disables all
        # VBA at load; execution is only requested after the parent completed a
        # static trust gate and therefore uses the dedicated worker's temporary
        # low-security open. Restore the setting immediately after Open.
        previous_security = excel.AutomationSecurity
        security_mode = str(job.get("security_mode", "disable"))
        excel.AutomationSecurity = 1 if security_mode == "trusted_execute" else 3
        try:
            book = excel.Workbooks.Open(
                workbook,
                ReadOnly=bool(job.get("read_only", False)),
                UpdateLinks=0,
                AddToMru=False,
                Notify=False,
                IgnoreReadOnlyRecommended=True,
            )
        finally:
            excel.AutomationSecurity = previous_security

        try:
            book.VBProject
        except Exception as exc:
            raise RuntimeError(
                "Cannot reach the Excel VBA project. Enable 'Trust access to the "
                "VBA project object model' in Excel's Trust Center (see the skill's "
                "references/setup.md)."
            ) from exc

        payload = _JOBS[job["action"]](excel, book, job)
        print(
            json.dumps(
                {
                    "ok": True,
                    "job_id": job_id,
                    "excel_pid": excel_pid,
                    "office_version": str(excel.Version),
                    "data": payload,
                },
                default=str,
            )
        )
        return 0
    except Exception as exc:
        print(
            json.dumps(
                {
                    "ok": False,
                    "job_id": job_id,
                    "excel_pid": excel_pid,
                    "office_version": str(excel.Version) if excel is not None else None,
                    "error_code": "excel_worker_error",
                    "error": str(exc),
                    "type": type(exc).__name__,
                    "traceback": traceback.format_exc(),
                },
                default=str,
            )
        )
        return 1
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    sys.exit(main())
