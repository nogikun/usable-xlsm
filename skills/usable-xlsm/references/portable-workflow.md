# macOS / Linux / no-Excel workflow

## Static development

Select the target OS once, independently of the agent's host. Run `doctor` or
`doctor --mode static`. Excel unavailability is expected and does not stop work.

```text
uv run --project skills/usable-xlsm usable-xlsm extract --workbook book.xlsm --output work/vba
uv run --project skills/usable-xlsm usable-xlsm preflight --workbook book.xlsm --operation inspect
uv run --project skills/usable-xlsm usable-xlsm plan --source work/vba --workbook book.xlsm
uv run --project skills/usable-xlsm usable-xlsm check --source work/vba --target macos
```

`preflight --operation edit/execute` evaluates the same trust/signature policy
statically, but passing it does not provide an execution backend or permission
to assume trust. Never add `--trust-workbook` merely to make the report green.
Keep a source folder under version control. Use individual-file checks while
editing and full checks for handoff. Transfer the workbook and changed source
files together; use `plan --partial` followed by the existing Windows
`update --partial` transaction. Never overwrite or recreate an existing workbook
with XlsxWriter: sheet data, drawings, signatures and VBA designers may be lost.

## Mac Excel

Mac Excel can open VBA workbooks, but this package's automated worker uses
Windows COM. Do not attempt pywin32 or claim `update/test/run` support on macOS.
For a trusted workbook, use a separate copy in Mac Excel's Visual Basic editor:
import standard `.bas` modules; copy only the code body into existing document
modules; use Debug > Compile and run the relevant procedures/tests manually.
Keep VBA codenames aligned with the workbook, inspect the buttons and fallback
paths, and save to a new candidate. Record these as manual Mac checks. Use the
Windows worker for the automated atomic release procedure when required.

## LibreOffice / UNO (optional compatibility checks)

LibreOffice is not an Excel substitute for production verification. Do not
convert and save the original workbook through it. Use a disposable copy,
an isolated profile, and the same explicit trust decision before macro execution.
Do not install it automatically just because Excel is unavailable.

First find an interpreter whose UNO import actually succeeds:

```text
/path/to/python -c "import uno; print(uno.__file__)"
```

A pip/uv Python often lacks `uno`. Official LibreOffice builds may supply
`/opt/libreoffice*/program/python` on Linux or
`/Applications/LibreOffice.app/Contents/Resources/python` on macOS; verify that
path exists rather than assuming it does. Distro packages may use system
Python with the distro's `python3-uno` package instead. Do not mix a UNO binary
built for one Python version with another interpreter via arbitrary PYTHONPATH.

Start soffice **before** connecting with UNO; use a fresh, absolute profile URI
for each disposable session and listen only on localhost:

```text
soffice -env:UserInstallation=file:///absolute/path/to/disposable-profile --headless --norestore --nodefault --nofirststartwizard "--accept=socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext"
```

Connect with that UNO-capable interpreter using
`uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext`.
Retry connection briefly while soffice starts, with a bounded timeout. Import
or connection failure is a capability failure, not a failed VBA test. Stop only
the process/profile created for this session. Never attach to or terminate the
user's everyday LibreOffice session.

## Portability rules

- Write one statement per line. In particular, split `Case "x": Proc` into a
  `Case` line followed by the call on its own line. Apply this to generated
  dispatch/test code as well as application modules.
- For shared Excel/LibreOffice error constants, write a numeric literal and
  retain its original `vbObjectError + n` expression in a comment. The bundled
  assertion helper follows this rule without changing the error number.
- Provide explicit fallback behavior for Excel-specific features such as
  `FormatConditions`; do not suppress errors and report success. For example,
  use direct `Interior.Color` formatting when equivalent to the user's intent.
  Verify that the fallback actually ran on the selected runtime.
- Review COM/ActiveX, Windows API declarations, paths, encodings and 32/64-bit
  declarations. Guard platform code and exercise the fallback separately.
- If a LibreOffice macro returns an empty result without an error, do not call
  it a pass. A compile error anywhere in a library can be a cause; isolate
  modules/procedures, inspect their compile errors, and rerun a known sentinel
  function. Require an explicit result for each discovered test.

`check --target libreoffice` is an advisory scanner, not a compatibility
compiler. It may flag guarded code, and cannot prove that a fallback exists.

Official references:
- https://help.libreoffice.org/latest/en-US/text/sbasic/python/main0000.html
- https://books.libreoffice.org/en/CG248/CG24814-CalcMacros.html
