"""Opt-in smoke check of freshly generated fixtures; never accepts user workbooks.

prepare/verify use the project's Python; run uses the distro's UNO Python.
See references/portable-workflow.md for Windows/WSL commands and limitations.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import signal
import socket
import subprocess
import sys
import tempfile
import time
from zipfile import ZipFile


NAMES = ("generated.xlsm", "seed-reused.xlsm")
EXPECTED_TEXT = "日本語の実行成功"
EXPECTED_TESTS = {
    "Test_Add": "pass", "Test_Japanese": "pass", "Test_SelectCase": "pass",
    "Test_AssertSuccess": "pass", "Test_DeliberateFailure": "fail",
}
RUNTIME = '''Attribute VB_Name = "Runtime"
Option Explicit
Public Function Sentinel() As Long
    Sentinel = 42
End Function
Public Sub Smoke()
    Worksheets("Runtime").Range("C1").Value = Sentinel()
    Worksheets("Runtime").Range("C2").Value = "日本語の実行成功"
    Worksheets("Runtime").Range("C3").Formula = "=SUM(A2:B2)"
    Worksheets("Runtime").Range("C4").Interior.Color = RGB(12, 34, 56)
    Worksheets("Runtime").Range("C5").Value = "button-called"
End Sub
Public Sub Test_Add()
    AssertEqual 2 + 3, 5
End Sub
Public Sub Test_Japanese()
    AssertEqual "日本語", "日本語"
End Sub
Public Sub Test_SelectCase()
    Dim result As Long
    Select Case "ok"
        Case "ok"
            result = 42
        Case Else
            result = 0
    End Select
    AssertEqual result, 42
End Sub
Public Sub Test_AssertSuccess()
    AssertTrue True
    AssertFalse False
    AssertAlmostEqual 0.1 + 0.2, 0.3
End Sub
Public Sub Test_DeliberateFailure()
    AssertEqual 1, 2, "expected negative control"
End Sub
'''


def check(condition, message):
    if not condition:
        raise RuntimeError(message)


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write_json(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def prepare(root):
    from usable_xlsm.creation import create_workbook, extract_project
    from usable_xlsm.testing import assert_module_source, discover, generate_runner

    root.mkdir(parents=True)  # Refuse to overwrite any existing fixture directory.
    source = root / "vba"
    source.mkdir()
    sources = {"Runtime.bas": RUNTIME, "UsableXlsmAssert.bas": assert_module_source()}
    sources["UsableXlsmRunner.bas"] = generate_runner(discover(sources))
    for name, text in sources.items():
        (source / name).write_text(text, encoding="utf-8")
    spec = root / "workbook.json"
    write_json(spec, {"sheets": [{"name": "Runtime", "rows": [["A", "B"], [20, 22]],
        "buttons": [{"cell": "E2", "macro": "Runtime.Smoke", "caption": "Run smoke"}]}]})
    create_workbook(spec, root / NAMES[0], source=source, experimental=True, target_os="linux")
    extract_project(root / NAMES[0], root / "vbaProject.bin")
    create_workbook(spec, root / NAMES[1], vba_project=root / "vbaProject.bin", target_os="linux")
    write_json(root / "inputs.json", {name: sha(root / name) for name in NAMES})


def verify(root):
    from usable_xlsm.core import _normalized_module_source, extract_vba

    report = json.loads((root / "runtime.json").read_text(encoding="utf-8"))
    inputs = json.loads((root / "inputs.json").read_text(encoding="utf-8"))
    check(set(inputs) == set(NAMES), "Unexpected fixture manifest")
    check([entry["name"] for entry in report["fixtures"]] == list(NAMES), "Missing runtime results")
    for entry in report["fixtures"]:
        original = root / entry["name"]
        saved = root / (original.stem + ".saved.xlsm")
        check(entry["ok"] and sha(original) == inputs[entry["name"]], "Runtime failed or input changed")
        check(sha(saved) == entry["saved_sha256"], "Saved candidate changed after runtime check")
        before, after = extract_vba(original), extract_vba(saved)
        entry["normalized_source_preserved"] = bool(before) and set(before) == set(after) and all(
            _normalized_module_source(before[name]) == _normalized_module_source(after[name]) for name in before)
        with ZipFile(original) as a, ZipFile(saved) as b:
            entry["vba_binary_preserved"] = a.read("xl/vbaProject.bin") == b.read("xl/vbaProject.bin")
    report["ok"] = report["ok"] and all(entry["normalized_source_preserved"] for entry in report["fixtures"])
    write_json(root / "verification.json", report)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    check(report["ok"], "VBA source changed during LibreOffice save")


def worker(root):
    import uno

    def prop(name, value):
        result = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        result.Name, result.Value = name, value
        return result

    def invoke(doc, library, module, procedure):
        uri = f"vnd.sun.star.script:{library}.{module}.{procedure}?language=Basic&location=document"
        return doc.getScriptProvider().getScript(uri).invoke((), (), ())[0]

    def check_runner(doc, library):
        raw = invoke(doc, library, "UsableXlsmRunner", "UsableXlsmRunAll")
        check(isinstance(raw, str) and raw.strip(), "Empty runner result")
        rows = [line.split("\t") for line in raw.splitlines() if line.strip()]
        check(len(rows) == 5 and all(len(row) == 5 for row in rows), "Incomplete runner results")
        check({row[1]: row[2] for row in rows} == EXPECTED_TESTS, raw)
        check("expected negative control" in raw, "Negative control did not fail as expected")
        return raw

    def check_button(doc, sheet):
        controls = [sheet.DrawPage.getByIndex(i) for i in range(sheet.DrawPage.Count)
            if sheet.DrawPage.getByIndex(i).supportsService("com.sun.star.drawing.ControlShape")]
        check(len(controls) == 1, "Form Control missing")
        events = sheet.DrawPage.Forms.getByIndex(0).getScriptEvents(0)
        action = next(event for event in events if event.EventMethod == "actionPerformed")
        check("Runtime.Smoke?" in action.ScriptCode, "Button lost its macro")
        sheet.getCellRangeByName("C5").String = ""
        doc.getScriptProvider().getScript(action.ScriptCode).invoke((), (), ())
        check(sheet.getCellRangeByName("C5").String == "button-called", "Button macro failed")

    inputs = json.loads((root / "inputs.json").read_text(encoding="utf-8"))
    check(set(inputs) == set(NAMES), "Run prepare first; only generated fixtures are supported")
    check(not (root / "runtime.json").exists(), "Use a fresh directory for each runtime check")
    report = {"version": subprocess.check_output(["soffice", "--version"], text=True).strip(),
              "fixtures": [], "ok": False}
    with tempfile.TemporaryDirectory(prefix="usable-xlsm-lo-") as temp:
        with socket.socket() as listener:
            listener.bind(("127.0.0.1", 0))
            port = listener.getsockname()[1]
        with (root / "soffice.log").open("w", encoding="utf-8") as log:
            process = subprocess.Popen(["soffice", f"-env:UserInstallation={(Path(temp) / 'profile').as_uri()}",
                "--headless", "--norestore", "--nodefault", "--nofirststartwizard",
                f"--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"],
                stdout=log, stderr=log, env={**os.environ, "SAL_USE_VCLPLUGIN": "svp"})
            desktop = None
            try:
                local = uno.getComponentContext()
                resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
                deadline = time.monotonic() + 20
                while True:
                    try:
                        context = resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
                        break
                    except Exception:
                        if time.monotonic() >= deadline or process.poll() is not None:
                            raise
                        time.sleep(0.2)
                desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
                for name in NAMES:
                    original = root / name
                    check(sha(original) == inputs[name], "Fixture changed since preparation")
                    copy = Path(temp) / name
                    shutil.copyfile(original, copy)
                    # Trust applies only to the fixture code above and the bundled test harness.
                    options = (prop("Hidden", True), prop("UpdateDocMode", 0), prop("MacroExecutionMode", 4))
                    doc = desktop.loadComponentFromURL(copy.as_uri(), "_blank", 0, options)
                    check(doc is not None, "Fixture did not load")
                    entry = {"name": name, "ok": False}
                    report["fixtures"].append(entry)
                    saved = root / (original.stem + ".saved.xlsm")
                    try:
                        library = next(lib for lib in doc.BasicLibraries.getElementNames()
                            if doc.BasicLibraries.getByName(lib).hasByName("Runtime"))
                        check(invoke(doc, library, "Runtime", "Sentinel") == 42, "Sentinel failed")
                        invoke(doc, library, "Runtime", "Smoke")
                        doc.calculateAll()
                        sheet = doc.Sheets.getByName("Runtime")
                        check(sheet.getCellRangeByName("C1").Value == 42, "Numeric result failed")
                        check(sheet.getCellRangeByName("C2").String == EXPECTED_TEXT, "Japanese result failed")
                        check(sheet.getCellRangeByName("C3").Value == 42, "Formula failed")
                        check(sheet.getCellRangeByName("C4").CellBackColor == 0x0C2238, "Formatting failed")
                        entry["runner_output"] = check_runner(doc, library)
                        check_button(doc, sheet)
                        doc.storeAsURL(saved.as_uri(), (prop("FilterName", "Calc MS Excel 2007 VBA XML"), prop("Overwrite", False)))
                        entry["saved_sha256"] = sha(saved)
                    finally:
                        doc.close(True)
                        check(sha(original) == inputs[name], "Original fixture changed")
                    doc = desktop.loadComponentFromURL(saved.as_uri(), "_blank", 0, options + (prop("ReadOnly", True),))
                    check(doc is not None, "Saved fixture did not reopen")
                    try:
                        check(invoke(doc, library, "Runtime", "Sentinel") == 42, "Saved sentinel failed")
                        check_runner(doc, library)
                        sheet = doc.Sheets.getByName("Runtime")
                        doc.calculateAll()
                        check(sheet.getCellRangeByName("C2").String == EXPECTED_TEXT, "Saved cell changed")
                        check(sheet.getCellRangeByName("C3").Value == 42, "Saved formula changed")
                        check(sheet.getCellRangeByName("C4").CellBackColor == 0x0C2238, "Saved formatting changed")
                        check_button(doc, sheet)
                        # A self-equality test alone passes even when both literals become '?'.
                        invoke(doc, library, "Runtime", "Smoke")
                        entry["saved_japanese_result"] = sheet.getCellRangeByName("C2").String
                        check(entry["saved_japanese_result"] == EXPECTED_TEXT, "Saved VBA corrupted Japanese text")
                        entry["ok"] = True
                    finally:
                        doc.close(True)
                report["ok"] = True
            except Exception as exc:
                report["error"] = f"{type(exc).__name__}: {exc}"
                raise
            finally:
                write_json(root / "runtime.json", report)
                if desktop is not None:
                    desktop.terminate()
                process.wait(timeout=5)


def run(root):
    check(os.name == "posix", "Run UNO through the Linux/macOS system Python (WSL on Windows)")
    process = subprocess.Popen([sys.executable, str(Path(__file__).resolve()), "worker", str(root)], start_new_session=True)
    try:
        check(process.wait(timeout=90) == 0, "LibreOffice runtime failed; candidate must not be accepted")
    finally:
        # All children inherit this new group; never kill another user's Office process.
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        process.wait()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("prepare", "run", "verify", "worker"))
    parser.add_argument("directory", type=Path)
    args = parser.parse_args()
    globals()[args.action](args.directory.resolve())
