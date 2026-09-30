"""Create new XLSM candidates from a declarative sheet/button specification."""

from io import BytesIO
import json
from pathlib import Path
import re
from zipfile import ZipFile

from oletools.olevba import VBA_Parser
import xlsxwriter

from .vba_project import project_bytes, write_new


def extract_project(workbook: str | Path, output: str | Path) -> Path:
    """Copy an existing project without opening or running its workbook."""
    with ZipFile(workbook) as archive:
        payload = archive.read("xl/vbaProject.bin")
        if any("vbaprojectsignature" in name.lower() or name.lower().startswith("_xmlsignatures/") for name in archive.namelist()):
            raise ValueError("Signed project: use your approved signing flow; this extractor does not preserve signatures")
    return write_new(output, payload)


def create_workbook(spec_path: str | Path, output: str | Path, *, vba_project: str | Path | None = None, source: str | Path | None = None, experimental: bool = False, target_os: str = "windows") -> dict:
    if (vba_project is None) == (source is None):
        raise ValueError("Choose exactly one of --vba-project or --source")
    if source is not None and not experimental:
        raise ValueError("Building a new binary project requires --experimental; Excel runtime validation remains pending")
    destination = Path(output)
    if destination.suffix.lower() != ".xlsm":
        raise ValueError("New workbook output must use .xlsm")
    if destination.exists():
        raise FileExistsError(f"Refusing to overwrite {destination}")
    spec = json.loads(Path(spec_path).read_text(encoding="utf-8"))
    sheets = spec.get("sheets", [{"name": "Sheet1"}])
    if not isinstance(sheets, list) or not sheets:
        raise ValueError("spec.sheets must be a nonempty array")
    workbook_name = spec.get("vba_name", "ThisWorkbook")
    document_names = [workbook_name] + [sheet.get("vba_name", f"Sheet{i}") for i, sheet in enumerate(sheets, 1)]
    if len(set(name.casefold() for name in document_names)) != len(document_names):
        raise ValueError("Workbook and worksheet VBA codenames must be unique")
    if any(not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{0,30}", name) for name in document_names):
        raise ValueError("VBA codenames must be ASCII identifiers of at most 31 characters")
    if source is not None:
        payload, _ = project_bytes(source, documents=document_names, target_os=target_os)
    else:
        payload = Path(vba_project).read_bytes()
    parser = VBA_Parser("vbaProject.bin", data=payload, relaxed=False)
    try:
        modules = {name: code for _, _, name, code in parser.extract_macros()}
    finally:
        parser.close()
    if not modules:
        raise ValueError("VBA project contains no extractable modules")
    missing = set(document_names) - {Path(name).stem for name in modules if name.endswith(".cls")}
    if missing:
        raise ValueError(f"VBA project lacks document codenames {sorted(missing)}; match spec.vba_name and sheet.vba_name to the template")

    # Produce in memory, then create the destination exclusively after ZIP checks.
    buffer = BytesIO()
    with xlsxwriter.Workbook(buffer, {"in_memory": True, "strings_to_formulas": False, "strings_to_urls": False}) as book:
        book.set_vba_name(workbook_name)
        if book.add_vba_project(BytesIO(payload), is_stream=True) != 0:
            raise ValueError("Unable to embed VBA project")
        for index, item in enumerate(sheets, 1):
            sheet = book.add_worksheet(item.get("name", f"Sheet{index}"))
            sheet.set_vba_name(document_names[index])
            for row, values in enumerate(item.get("rows", [])):
                if sheet.write_row(row, 0, values) != 0:
                    raise ValueError("Sheet row exceeds worksheet limits")
            for cell, value in item.get("cells", {}).items():
                if sheet.write(cell, value) != 0:
                    raise ValueError(f"Invalid worksheet cell: {cell}")
            for column, width in item.get("column_widths", {}).items():
                sheet.set_column(column, width)
            for button in item.get("buttons", []):
                macro = button["macro"]
                if not re.fullmatch(r"(?:[A-Za-z]\w*\.)?[A-Za-z]\w*", macro):
                    raise ValueError(f"Invalid button macro name: {macro}")
                if sheet.insert_button(button["cell"], {key: value for key, value in button.items() if key != "cell"}) != 0:
                    raise ValueError("Button exceeds worksheet limits")
    workbook_bytes = buffer.getvalue()
    with ZipFile(BytesIO(workbook_bytes)) as archive:
        if archive.testzip() or archive.read("xl/vbaProject.bin") != payload:
            raise ValueError("New workbook failed ZIP/project integrity check")
    write_new(destination, workbook_bytes)
    return {"ok": True, "output": str(destination), "modules": sorted(modules), "sheets": len(sheets), "experimental": source is not None, "runtime_verified": False, "validation_pending": ["Excel open/compile", "VBA tests", "button actions", f"target OS: {target_os}"]}
