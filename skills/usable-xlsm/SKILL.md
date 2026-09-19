---
name: usable-xlsm
description: Safely inspect, edit, test, and release VBA and workbook-object changes in trusted .xlsm/.xlsb/.xltm workbooks on a dedicated Windows Excel worker. Use this skill instead of xlsx when VBA, signatures, workbook objects, or preservation-sensitive content is involved or unknown. Use xlsx for ordinary .xlsx/.csv/.tsv work and verified low-risk cell-only edits. Do not use for arbitrary untrusted Office files or server-side concurrent Office automation.
---

# usable-xlsm

Treat a macro-enabled workbook as executable code. Static reading is safe;
opening or executing it through Excel requires an explicit trust decision.

## Choose the shortest supported route

`xlsx` also triggers on `.xlsm`, so route by the operation, not by the file
extension alone. When both skills could apply, `usable-xlsm` takes precedence
whenever VBA, signatures, workbook objects, external links, or preservation
requirements exist or are unknown.

| Input and task | Route |
|---|---|
| `.xlsx`, `.csv`, `.tsv`, ordinary cells/formulas/formatting | `xlsx` |
| `.xlsm`/`.xltm`: add, replace, remove, run, or test VBA | `usable-xlsm` |
| `.xlsm`/`.xltm`: controls, UserForms, shapes, workbook/sheet objects, or signed content | `usable-xlsm` |
| `.xlsb`, any macro or cell edit | `usable-xlsm`; `openpyxl` is not an `.xlsb` editor |
| `.xlsm` cell-only edit with no VBA/object change | Use `xlsx` only after the preservation gate below; otherwise use `usable-xlsm` |

`openpyxl` with `keep_vba=True` preserves an existing VBA project while saving
some workbook edits; it does not create or import new VBA source modules, and
it does not prove that shapes, controls, UserForms, external links, or other
package parts survive. A saved workbook is a candidate until those parts have
been checked. Never use this as a shortcut for VBA injection.

For the `xlsx` route on an `.xlsm`, all of these must be true before saving a
candidate copy: the file is `.xlsm` or `.xltm` (not `.xlsb`), it is unsigned,
there are no controls, shapes, UserForms, external links, or other
preservation-sensitive parts, and the requested change is cells/formulas/
formatting only. Compare the VBA and package/object inventory after saving and
follow the `xlsx` skill's formula-recalculation rules. If any item is unknown,
fail closed and use this skill; the current CLI has no general-purpose cell
edit command that makes `openpyxl` safe for such a workbook.

## Non-negotiable production rules

- Never call Excel COM directly. Use `usable-xlsm`; its worker is disposable,
  detects blocking VBA dialogs, and only terminates the Excel PID it owns.
- Run on a dedicated Windows account or disposable VM. One worker runs one
  Excel job at a time, and production commands refuse to start while another
  Excel process is present.
- Never infer that a workbook is trusted because it is local or has no Mark of
  the Web. Every command that edits or executes must receive an approved TOML
  policy or an explicit `--trust-workbook` attestation for that exact input.
  A previous `preflight` result is not an authorization token for a later
  `update`, `test`, or `run` command.
- Never remove MOTW, weaken Trust Center policy, or add a Trusted Location on
  the user's behalf.
- Never edit the original through COM. Updates go to a staging copy, are saved,
  closed, re-extracted, compared, tested on fresh copies when requested, and
  only then replace the supplied target atomically. A managed backup and JSONL
  audit record are created.
- Signed workbooks are blocked for editing unless signature invalidation is
  explicitly acknowledged. A promoted result is an unsigned candidate until
  an external signing process re-signs and independently verifies it.
- `--no-test` is a development shortcut, not a production claim. It still
  atomically replaces the workbook passed as `--workbook`, so use it only with
  a disposable development copy.

Read [references/security.md](references/security.md) before accepting or
executing files from a new source. Read
[references/production.md](references/production.md) when provisioning a
worker, CI job, signing flow, or monitoring.

## Generic VBA scope

This skill is a generic VBA source workflow, not a calculator workflow. Treat
the requested module names, procedures, document modules, and workbook
objects as inputs; the calculator in `sample/` is only an end-to-end fixture.

- Standard `.bas` modules can be added, replaced, or removed with `sync`.
- Editable `.cls` class modules can be added, replaced, or removed with `sync`.
- Existing worksheet and `ThisWorkbook` document modules can be updated, but
  they are not removable or creatable by source-file injection.
- `.frm` UserForms are extractable for inspection but read-only because the
  `.frx` binary designer must remain paired with the form.
- ActiveX, signed VBA, and XLM changes remain explicit advanced operations;
  do not silently convert them into a source-overlay task.

The requested source files are the patch, not a mandatory replacement of the
whole project unless `--sync` is explicitly selected. Never invent procedure
names or generate arbitrary VBA when the user has not supplied the source or
asked for generation.

## Choose the verification budget

Repository test files are developer QA; a field operation does not run the
entire Python test suite. Workbook `Test_*` procedures are a separate runtime
test layer. Use the smallest check that proves the current change, then reserve
the full suite for release:

| Situation | Route | What it proves |
|---|---|---|
| Every development iteration | `apply --no-test` on a disposable copy | live Excel import, save/close, VBA re-extraction, object/package verification |
| A behavior change with one known entry point | add one `run --macro Module.Procedure` smoke call on the disposable result | the changed runtime path works in real Excel |
| A workbook with `Test_*` procedures during development | `test --filter Test_Name --shared-copy` | one targeted runtime test without starting Excel once per test |
| Release or handoff | `apply` with default isolated tests | all discovered tests on fresh copies |

Do not run the full repository or workbook suite after every source edit.
Development is still real: `apply` uses Excel and the targeted smoke/test
uses Excel; only unrelated tests are deferred. If no suitable smoke entry
point or `Test_*` procedure exists, report that runtime behavior remains
unproven rather than compensating with a large unrelated test run.

## Fast VBA development loop

Use this loop when iterating on VBA in a disposable copy. It removes repeated
test jobs, while retaining trust checks, syntax checks, staging, re-extraction,
backup, atomic replacement, and audit logging.

1. Check the worker once per session. `doctor` reports readiness; it does not
   cache authorization or replace the per-operation preflight.

   ```powershell
   uv run --project skills/usable-xlsm usable-xlsm doctor
   ```

2. Create a disposable workbook target and extract to a **new empty** source
   directory. Do not reuse a directory that may contain stale modules.

   ```powershell
   Copy-Item sample/example.xlsm work/example.dev.xlsm
   New-Item -ItemType Directory -Force work/vba-dev | Out-Null
   uv run --project skills/usable-xlsm usable-xlsm extract `
     --workbook work/example.dev.xlsm --output work/vba-dev
   ```

3. Edit `.bas` and editable `.cls` files as ordinary source, then run the
   optional early syntax check. `update` repeats authorization, module-name,
   and syntax checks, so omit these diagnostics when the loop is already
   stable and the absolute shortest command sequence is needed.

   ```powershell
   uv run --project skills/usable-xlsm usable-xlsm check --source work/vba-dev
   ```

   Existing `.frm` source may be inspected but must not be modified: its
   paired binary designer cannot be safely round-tripped.

4. Promote the development iteration with tests disabled. Supply either the
   approved policy on **this command** or an explicit attestation for this
   exact disposable workbook.

   ```powershell
   uv run --project skills/usable-xlsm usable-xlsm update `
     --source work/vba-dev --workbook work/example.dev.xlsm `
     --policy usable-xlsm.toml --no-test
   ```

   `--trust-workbook` may replace `--policy` only after the operator has
   explicitly attested to that exact input. Do not copy this command to a
   production target merely because it is fast.

5. If the change affects runtime behavior, run one targeted smoke entry point
   on the disposable result. `run` uses a copy by default, so this proves the
   real Excel path without dirtying the development workbook:

   ```powershell
   uv run --project skills/usable-xlsm usable-xlsm run `
     --workbook work/example.dev.xlsm --macro Module1.Smoke `
     --policy usable-xlsm.toml
   ```

   Replace `Module1.Smoke` with the user-supplied public procedure. If the
   workbook has a focused `Test_*`, use `test --filter Name --shared-copy`
   instead. Do not run every test as a substitute for choosing a runtime
   entry point.

### Adding or removing modules

Without `--sync`, the source directory's filenames must exactly match the
workbook's extracted modules. This is the safest path for changing an existing
module.

Use `--sync` only when the freshly extracted directory has been reviewed as the
complete desired set. Sync can add or remove standard `.bas` and class `.cls`
modules; an absent removable module means deletion. Document modules such as
`Sheet1.cls` and `ThisWorkbook.cls` are update-only, and UserForms (`.frm` plus
their binary designer) are read-only. Review the returned `added`, `updated`,
and `removed` manifest before accepting the result.

```powershell
uv run --project skills/usable-xlsm usable-xlsm update `
  --source work/vba-dev --workbook work/example.dev.xlsm `
  --policy usable-xlsm.toml --sync --no-test
```

The short loop proves static syntax, Excel's staged import/re-extraction, and
the declared workbook-object/package changes. It does not prove arbitrary
runtime behavior or full-project compilation; use the targeted smoke step
above for that. `verify` below is cleanup verification only.

## One-shot script: the shortest supported edit path

For repeated VBA and Form Control edits, use one JSON script. `apply` is the
normal entry point: it extracts the current modules into a disposable
directory, overlays the listed `.bas`/`.cls` files, applies the buttons in the
same Excel job, verifies the closed saved package, and promotes the result
through the normal staging/backup/audit path.

The script contract is strict. The allowed top-level keys are `modules`,
`remove_modules`, `worksheets`, `cells`, and `buttons`. Unknown keys fail
before Excel starts; this prevents a misspelled or future operation from being
silently reported as successful.

```json
{
  "modules": ["vba/Module1.bas", "vba/NewModule.bas"],
  "worksheets": [{"name": "Calculator", "create": true, "visible": "visible"}],
  "cells": [
    {"sheet": "Calculator", "address": "B2", "value": 0, "number_format": "0.00"},
    {"sheet": "Calculator", "address": "C2", "formula": "=B2+1"}
  ],
  "buttons": [
    {
      "sheet": "Sheet1",
      "name": "btnRun",
      "caption": "Run",
      "macro": "Module1.Run",
      "left": 12,
      "top": 24,
      "width": 120,
      "height": 24,
      "replace": true
    }
  ]
}
```

Run it on a disposable target during development:

```powershell
uv run --project skills/usable-xlsm usable-xlsm apply `
  --script changes.json --workbook work/example.dev.xlsm `
  --policy usable-xlsm.toml --no-test
```

Script module paths are relative to the JSON file. Existing modules are kept;
listed files replace matching names and new names are added. `remove_modules`
may be added for intentional standard `.bas`/class `.cls` removal; document
modules such as `Sheet1.cls` and `ThisWorkbook.cls` are protected and are not
deletion targets. The script automatically uses sync semantics for a module
change and reports `updated`, `added`, and `removed`. It cannot edit UserForms.
Worksheet entries are idempotent: an existing sheet is reused, a missing sheet
is created only when `create: true`, and `visible` is one of `visible`,
`hidden`, or `veryHidden`. Cell entries target one A1 cell, accept a scalar
value (including `null` to clear it) or an Excel formula, and reject duplicate
addresses. Cell writes replace existing contents by default; set
`overwrite: false` to fail when the target is non-empty. `number_format` is
optional. Worksheet and cell changes are applied before buttons in the same
Excel job.

Button names must be unique per sheet; `replace: true` is required to replace
an existing button. Button assignment does not run the macro while applying
the layout. The result includes
`buttons_verified` from the live worker and `buttons_saved_verified` from the
closed, saved OOXML package. It also reports `test_mode` (`null`, `isolated`,
or `shared_copy`), `excel_jobs`, `worksheets_saved_verified`, and
`cells_saved_verified` so performance and persistence results are reproducible.

The button macro is an assignment, not an execution. Use a trusted test suite
or an explicit `run` on a disposable copy to prove behavior. A missing target
macro or a VBA runtime error is not proven by button-layout verification alone.

Do not add a calculator-specific command, a layout engine, UserForm/ActiveX
support, or arbitrary VBA generation to this path. Add a new declarative field
only after its conflict rules, saved-package verification, and timing impact
are specified.

If the source directory already exists, `update --buttons buttons.json` can
apply the same button manifest alongside the existing VBA update. Use `apply`
when the goal is one self-contained repeatable change set.

## Tested release loop

Measure the real Windows/Excel path with disposable copies. Use at least three
runs and report the p50 wall time; do not compare an `apply` run with a
standalone `test` run as if they were the same operation:

```powershell
uv run --project skills/usable-xlsm python skills/usable-xlsm/scripts/benchmark_apply.py `
  --workbook sample/example2.xlsm --script work/benchmark/apply.json `
  --output-dir work/benchmark/runs --mode dev --runs 3
```

Use `--mode isolated` for the release default and `--mode shared` only when
the shared-copy exception is explicitly acceptable. Compare the same mode,
same script, and same workbook across revisions.

Run the default tested update against the intended release target only after
the development copy is satisfactory:

```powershell
uv run --project skills/usable-xlsm usable-xlsm update `
  --source work/vba-release --workbook release.xlsm `
  --policy usable-xlsm.toml
uv run --project skills/usable-xlsm usable-xlsm verify --target release.xlsm
```

The default requires at least one `Public Sub Test_*()` and runs each test on a
fresh workbook copy. For `sample/example.xlsm`, the extracted sample has no
discoverable `Test_*` procedure, so a tested release requires adding a test to
the source first. `--allow-no-tests` is an explicit untested exception, not an
equivalent release path. `--shared-test-copy`, `--no-test`, and
`--allow-signature-removal` are also policy exceptions; never add them silently.

`verify` checks for surviving harness modules, scratch copies, unreadable
workbooks, and Excel processes. It does not prove VBA behavior, formulas,
signatures, workbook objects, or release correctness. The update itself already
performs authorization, module validation, syntax checking, staged
re-extraction, and atomic promotion; standalone `preflight` and `check` are
optional early diagnostics, not a way to skip those checks.

Restore the newest managed backup without launching Excel:

```powershell
uv run --project skills/usable-xlsm usable-xlsm restore --workbook release.xlsm
```

Create a redacted diagnostic bundle when escalation is needed:

```powershell
uv run --project skills/usable-xlsm usable-xlsm support-bundle `
  --workbook release.xlsm --output support.json
```

The bundle contains versions, process IDs, scratch filenames, and recent
redacted audit events—not workbook bytes, VBA source, arguments, or cell data.

## Workbook-object finalizers

When an update must create or modify workbook objects not represented by
`.bas`/`.cls` source, run a deliberately chosen trusted finalizer on the
staging copy as part of the same transaction:

```powershell
uv run --project skills/usable-xlsm usable-xlsm update `
  --source work/vba-release --workbook release.xlsm `
  --policy usable-xlsm.toml --post-macro Module1.InstallControlButtons
```

`--post-macro` runs after VBA source is applied and before static verification,
isolated tests, and atomic promotion. Use it only for a trusted, deterministic
finalizer. Its side effects on the staging copy become part of the promoted
workbook. Repeated macro arguments can be supplied with `--post-macro-arg`.

## Inspecting and executing without promotion

Static inspection does not launch Excel:

```powershell
uv run --project skills/usable-xlsm usable-xlsm preflight `
  --workbook book.xlsm --operation inspect
uv run --project skills/usable-xlsm usable-xlsm extract `
  --workbook book.xlsm --output work/vba
uv run --project skills/usable-xlsm usable-xlsm check --source work/vba
```

For edit or execute preflight, pass the same policy or explicit attestation
that will be used by the subsequent command:

```powershell
uv run --project skills/usable-xlsm usable-xlsm preflight `
  --workbook book.xlsm --operation edit --policy usable-xlsm.toml
```

Run a macro on a disposable copy by default:

```powershell
uv run --project skills/usable-xlsm usable-xlsm test `
  --workbook book.xlsm --policy usable-xlsm.toml
uv run --project skills/usable-xlsm usable-xlsm run `
  --workbook book.xlsm --macro Module1.MyMacro --policy usable-xlsm.toml
```

`run --in-place` is an explicitly approved interactive exception. Prefer
`update --post-macro` when the macro is a deterministic part of the release.

## VBA test convention

Tests live in standard `.bas` modules:

```vb
Public Sub Test_AddTwo()
    AssertEqual AddTwo(2, 3), 5
End Sub
```

Optional public, argument-free `TestSetup` and `TestTeardown` procedures run
around each test. Assertion helpers are injected only into disposable copies.

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
