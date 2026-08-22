# VBA and implementation notes

## Architecture

```text
CLI
  -> security.py     static trust gate; never opens Excel
  -> core.py         locks, staging, backup, verification, promotion, audit
     -> syntax.py    ANTLR VBA syntax parse
     -> testing.py   test discovery and runner generation
     -> runner.py    exclusive subprocess watchdog and PID ownership
        -> _worker.py  one disposable Excel COM instance
```

The worker publishes a JSON handshake containing the job ID, Python worker PID,
and Excel PID before opening a workbook. The watchdog kills only that exact
Excel PID. If ownership is unknown, it reports an error and requires host
recycle instead of guessing.

## Modal dialogs

VBA compile and runtime errors can raise Win32 modal dialogs that
`DisplayAlerts=False` does not suppress. A hidden Excel process then blocks its
COM caller indefinitely.

The watchdog polls dialogs owned by its Excel PID, extracts their text, and
terminates that owned process tree. Never bypass it with direct
`win32com.client.Dispatch`.

## Macro security

`Application.AutomationSecurity` is set before every `Workbooks.Open`:

- static editing and COM inspection use ForceDisable;
- execution occurs only after static authorization, on a dedicated worker, and
  uses the explicit trusted-execution mode for the open;
- the prior setting is restored immediately after opening;
- events are disabled unless requested;
- links are not updated.

XLM macros are rejected in preflight because ForceDisable does not cover them.

## Lazy VBA compilation

VBA compiles a called procedure and its dependencies rather than reliably
compiling the entire project through COM. Therefore:

- the static grammar check covers every exported module's structure;
- generated test calls are static so each test and its dependencies compile;
- semantic errors in uncalled procedures can remain undetected;
- release tests should exercise every supported entry point.

Do not claim a full-project compile unless a controlled visible VBE compile step
has actually been performed and observed.

## Source normalization

`olevba` exports component `Attribute` records that cannot all be reinserted as
ordinary code. The writer strips `Attribute ...` lines; Excel regenerates
component attributes. Verification applies the same normalization and compares
all requested module bodies after the staged workbook has been saved and
closed.

VBA source uses CRLF. Writers pass `newline=""` to avoid doubled carriage
returns on Windows.

## Component types

| Type | Meaning | Production behavior |
|---|---|---|
| 1 | Standard module (`.bas`) | May be updated, added, or removed with sync. |
| 2 | Class module (`.cls`) | May be updated, added, or removed with sync. |
| 3 | UserForm (`.frm` + `.frx`) | Read-only; changed or new form source is rejected. |
| 100 | Workbook/sheet document module | May be updated; never added or removed. |

UserForms have a binary designer. Updating only the text portion can corrupt or
silently desynchronize the form, so this implementation rejects such changes.

## Test harness

A test is a public, argument-free `Sub` named `Test_*` in a standard module.
Optional public `TestSetup` and `TestTeardown` procedures are scoped to that
module.

The generated runner uses `On Error Resume Next` so assertion and runtime
errors become data instead of modal dialogs. It saves the setup/test error,
runs teardown, and records teardown failure before clearing `Err`. A teardown
failure therefore cannot produce a false green.

Production defaults run every test on a fresh workbook copy. This prevents one
case's cell or workbook mutations from contaminating the next case. The
assertion and runner modules are injected only into those copies.

The following are failures:

- assertion or runtime error;
- setup or teardown error;
- a discovered test that never reports;
- zero tests when tests are required;
- scratch-copy deletion failure;
- surviving harness modules or staging artifacts.

## Atomic update

Excel only writes the staging copy. After it closes, the orchestrator:

1. re-extracts VBA and compares requested normalized sources;
2. verifies the editable component manifest for sync operations;
3. executes isolated tests;
4. confirms the original hash still equals the preflight hash;
5. uses `os.replace` on the same volume to promote the staging file.

Any earlier failure leaves the original unchanged. Managed backups remain
available for explicit `restore`.
