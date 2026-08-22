---
name: usable-xlsm
description: Safely inspect, edit, test, and release VBA and workbook-object changes in trusted .xlsm/.xlsb/.xltm workbooks on a dedicated Windows Excel worker. Use for VBA or macro-enabled workbook development, repair, review, regression testing, and controlled promotion. Do not use for plain .xlsx data work, arbitrary untrusted Office files, or server-side concurrent Office automation.
---

# usable-xlsm

Treat a macro-enabled workbook as executable code. Static reading is safe;
opening or executing it through Excel requires an explicit trust decision.

## Non-negotiable production rules

- Never call Excel COM directly. Use `usable-xlsm`; its worker is disposable,
  detects blocking VBA dialogs, and only terminates the Excel PID it owns.
- Run on a dedicated Windows account or disposable VM. One worker runs one
  Excel job at a time, and production commands refuse to start while another
  Excel process is present.
- Never infer that a workbook is trusted because it is local or has no Mark of
  the Web. Use an approved TOML policy, or `--trust-workbook` only after the
  user explicitly attests to that exact input.
- Never remove MOTW, weaken Trust Center policy, or add a Trusted Location on
  the user's behalf.
- Never edit the original through COM. Updates go to a staging copy, are saved,
  closed, re-extracted, compared, tested on fresh copies, and only then replace
  the original atomically. A managed backup and JSONL audit record are created.
- Signed workbooks are blocked for editing unless signature invalidation is
  explicitly acknowledged. Production release still requires an approved
  external re-signing step.

Read [references/security.md](references/security.md) before accepting or
executing files from a new source. Read
[references/production.md](references/production.md) when provisioning a
worker, CI job, signing flow, or monitoring.

## Standard workflow

Run from the repository root.

```powershell
uv run --project skills/usable-xlsm usable-xlsm doctor
```

The worker is ready only when this returns `"ok": true`. `AccessVBOM` belongs
only on the dedicated automation profile; see
[references/setup.md](references/setup.md).

Inspect without launching Excel:

```powershell
uv run --project skills/usable-xlsm usable-xlsm preflight `
  --workbook book.xlsm --operation edit --policy usable-xlsm.toml
uv run --project skills/usable-xlsm usable-xlsm extract `
  --workbook book.xlsm --output work/vba
uv run --project skills/usable-xlsm usable-xlsm check --source work/vba
```

Edit `.bas` and `.cls` files as ordinary source. Keep filenames equal to VBA
component names. Existing `.frm` source may be inspected but must not be
modified: its paired binary designer cannot be round-tripped safely.

Promote an update:

```powershell
uv run --project skills/usable-xlsm usable-xlsm update `
  --source work/vba --workbook book.xlsm --policy usable-xlsm.toml
```

When the update also needs to create or modify workbook objects that are not
represented by `.bas`/`.cls` source (for example Form Control buttons), run a
trusted finalizer on the staging copy as part of the same transaction:

```powershell
uv run --project skills/usable-xlsm usable-xlsm update `
  --source work/vba --workbook book.xlsm --policy usable-xlsm.toml `
  --post-macro Module1.InstallControlButtons
```

`--post-macro` runs after the VBA source is applied and before static
verification, isolated tests, and atomic promotion. The original workbook is
not opened for this step, so a separate `run --in-place` command is not needed.
Use it only for an explicitly chosen, trusted finalizer; its side effects on
the staging copy become part of the promoted workbook. Repeated macro
arguments can be supplied with `--post-macro-arg`.

By default this requires at least one `Public Sub Test_*()`, isolates every test
on a fresh workbook copy, and refuses promotion on any test, teardown, cleanup,
integrity, or audit failure. `--allow-no-tests`, `--shared-test-copy`,
`--no-test`, and `--allow-signature-removal` are release-policy exceptions;
never add them silently.

Run tests or a macro without changing the original:

```powershell
uv run --project skills/usable-xlsm usable-xlsm test `
  --workbook book.xlsm --policy usable-xlsm.toml
uv run --project skills/usable-xlsm usable-xlsm run `
  --workbook book.xlsm --macro Module1.MyMacro --policy usable-xlsm.toml
```

`run` uses a disposable copy by default. `--in-place` is an explicit exception
for an approved interactive case, not a normal development shortcut. Prefer
`update --post-macro` when the macro is a deterministic part of the workbook
release, because it keeps the finalizer inside the staged update transaction.

Confirm no harness or scratch artifacts remain:

```powershell
uv run --project skills/usable-xlsm usable-xlsm verify --target book.xlsm
```

Restore the newest managed backup:

```powershell
uv run --project skills/usable-xlsm usable-xlsm restore --workbook book.xlsm
```

Create a redacted diagnostic bundle when escalation is needed:

```powershell
uv run --project skills/usable-xlsm usable-xlsm support-bundle `
  --workbook book.xlsm --output support.json
```

The bundle contains versions, process IDs, scratch filenames, and recent
redacted audit events—not workbook bytes, VBA source, arguments, or cell data.

## VBA test convention

Tests live in standard `.bas` modules:

```vb
Public Sub Test_AddTwo()
    AssertEqual AddTwo(2, 3), 5
End Sub
```

Optional public, argument-free `TestSetup` and `TestTeardown` procedures run
around each test. A teardown error fails the case even when the test body
passed. Assertion helpers are injected only into disposable copies.

## Interpreting failures

- Exit `2`: trust/security preflight blocked the operation.
- Exit `1`: syntax, Excel, test, verification, cleanup, or audit failure.
- `excel_host_not_clean`: another Excel process exists; use a clean dedicated
  worker. Do not kill unidentified processes.
- `excel_pid_unknown`: recycle the worker host. The watchdog deliberately did
  not guess which Excel process to terminate.
- `no_tests`: add a discoverable `Test_*` suite or obtain explicit approval for
  `--allow-no-tests`.
- `cleanup_failed`: do not treat the run as successful; inspect with `verify`.

For VBA parsing, component types, lazy compilation, and harness internals, read
[references/vba-notes.md](references/vba-notes.md).
