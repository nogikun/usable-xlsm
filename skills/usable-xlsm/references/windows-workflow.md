# Windows Excel workflow

## Standard workflow

Run from the repository root.

```powershell
uv run --project skills/usable-xlsm usable-xlsm doctor
```

The worker is ready only when this returns `"ok": true`. `AccessVBOM` belongs
only on the dedicated automation profile; see
[setup.md](setup.md).

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

## Small module changes

For one new standard module, give the UTF-8 `.bas` file directly. Inspect the
addition without opening Excel, then apply it with the existing transaction:

```powershell
uv run --project skills/usable-xlsm usable-xlsm plan `
  --source work/NewModule.bas --workbook book.xlsm --add-only --diff
uv run --project skills/usable-xlsm usable-xlsm update `
  --source work/NewModule.bas --workbook book.xlsm --add-only --policy usable-xlsm.toml
```

`--add-only` accepts exactly one standard `.bas` file. Use a filename stem of
1–31 ASCII letters, digits or underscores, starting with a letter. An optional
`Attribute VB_Name` must match it. Any existing component name, ignoring case
and component type, blocks addition. This mode cannot use `--partial` or
`--sync`; it does not rewrite or remove existing modules. After saving, both
the new source and all previously extracted sources are compared, and the
default full isolated VBA tests still run before promotion.

For a small edit to existing modules, put only the changed files in a source
directory and use `--partial`. Opt into skipping an unchanged iteration:

```powershell
uv run --project skills/usable-xlsm usable-xlsm plan `
  --source work/patch --workbook book.xlsm --partial --diff
uv run --project skills/usable-xlsm usable-xlsm update `
  --source work/patch --workbook book.xlsm --partial --skip-unchanged --policy usable-xlsm.toml
```

`--skip-unchanged` still checks trust, source syntax, the workbook lock and its
input hash, and writes an audit record. If the compared source has no additions,
updates or deletions and no `--post-macro` is requested, it returns
`status: unchanged`, `excel_jobs: 0`, `runtime_verified: false`, `tests: null`.
It creates no backup, does not save the workbook and does not run runtime tests.
Source normalization ignores `Attribute` lines, line endings, trailing spaces
and trailing blank lines. This is a development shortcut, not release evidence;
omit `--skip-unchanged` when tests must run even without source changes.

The update result's `changes` object lists `added`, `changed`, `removed` and
`unchanged` source filenames. `plan` lists the same fields at the top level.
`--add-only` shows no deletions; `--partial` keeps omitted modules unchanged.

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
[vba-notes.md](vba-notes.md).
