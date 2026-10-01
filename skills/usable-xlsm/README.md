# usable-xlsm

Portable VBA development for Windows, macOS and Linux, with the existing
dedicated Windows Excel worker for staged updates and isolated runtime tests.

```text
uv sync --project skills/usable-xlsm --frozen
uv run --project skills/usable-xlsm usable-xlsm environment --target-os windows
uv run --project skills/usable-xlsm usable-xlsm doctor
uv run --project skills/usable-xlsm usable-xlsm --help
```

- Select the workbook user's OS once, separately from the execution host.
- Extract and edit source on any supported host; use `plan` for compact diffs,
  single-file `check` for quick feedback, and `--partial` for existing-module updates.
- Keep Windows COM watchdog, trust policy, staged verification, isolated full
  tests, signatures, backups, audits and atomic promotion.
- Use `create` for new sheet/button candidates with a known seed project.
  Source-only binary building is explicit experimental functionality.
- Use `check --target macos/libreoffice` for advisory portability findings.
  Static/structural checks always report `runtime_verified: false`.
- For optional LibreOffice save/reopen validation, use the generated-fixture
  smoke check in [the portable workflow](references/portable-workflow.md#repeatable-savereopen-smoke-check).
  It rejects Japanese VBA corruption and source loss without changing user workbooks.

Operational instructions live in `SKILL.md`. Load the Windows, portable or
creation reference only when the selected workflow requires it.
