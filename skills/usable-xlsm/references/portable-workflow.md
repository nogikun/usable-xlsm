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

### Repeatable save/reopen smoke check

Use `tests/libreoffice_smoke.py` only with the fresh fixtures it creates. It
does not accept a user's workbook. The isolated profile prevents interference
with an existing Office session; it is not a sandbox for untrusted macro code.
Review the script and bundled assertion/runner modules before running it.

From the repository root on Linux, with LibreOffice and the matching
`python3-uno` installed:

```text
uv run --project skills/usable-xlsm python skills/usable-xlsm/tests/libreoffice_smoke.py prepare work/lo-smoke-1
/usr/bin/python3 skills/usable-xlsm/tests/libreoffice_smoke.py run work/lo-smoke-1
uv run --project skills/usable-xlsm python skills/usable-xlsm/tests/libreoffice_smoke.py verify work/lo-smoke-1
```

On Windows, run `prepare` and `verify` with Windows uv. For the middle command,
use WSL's UNO Python with absolute `/mnt/c/...` paths to both the script and the
fixture directory, for example (adjust the checkout path):

```text
wsl.exe -d Ubuntu -- /usr/bin/python3 /mnt/c/Users/takah/Documents/git/usable-xlsm/skills/usable-xlsm/tests/libreoffice_smoke.py run /mnt/c/Users/takah/Documents/git/usable-xlsm/work/lo-smoke-1
```

Use a new directory for each attempt. `prepare` refuses an existing directory;
`run` checks input hashes, copies inputs to a temporary directory, disables
external-document updates, saves separate candidates, and reopens them read-only.
The supervisor allows 90 seconds and cleans up only its own process group.
All three commands must succeed. Inspect `verification.json` and `soffice.log`
on failure; never promote a failed candidate or overwrite the original.

The check covers source-built and reused-project candidates, a nonempty numeric
sentinel, Japanese text compared to a fixed Python value, formulas, direct cell
formatting, the registered Form Control macro, and four passing tests plus one
deliberately failing assertion. It reruns VBA after saving, compares module names
and normalized source, verifies saved-file hashes, and confirms inputs are unchanged.
The button check invokes its registered macro; it does not simulate a GUI click.

WSL Ubuntu 22.04 testing on 2026-10-01 found:

| LibreOffice | Initial runtime | Japanese VBA after XLSM save/reopen |
|---|---|---|
| 7.3.7.2 | Passed | Corrupted to `?`; rejected |
| 26.2.5.2 | Passed | Preserved; normalized source comparison passed |

On 7.3.7.2, setting Calc's `Filter/Import/VBA/UseExport` to false did not
prevent this corruption. On 26.2.5.2, default settings passed. Choose a version
that passes this check in your environment; these observations do not establish
a minimum supported version. For Ubuntu 22.04 this comparison used Calc and
`python3-uno` from the official `jammy-backports` repository.

Even on 26.2.5.2, `vbaProject.bin` bytes and some document-module attributes
changed. Normalized-source equality ignores `Attribute` lines, line endings,
trailing spaces and trailing blank lines; it does not prove preservation of
signatures, references, designers, arbitrary Excel features or binary identity.
The reused project here is extracted from the generated fixture, not an
Excel-authored seed. This check stays optional and separate from the portable
unit suite. Use the existing Windows Excel transaction for production release.

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
