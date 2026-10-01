---
name: usable-xlsm
description: Inspect, create, repair, and statically validate macro-enabled Excel workbooks and VBA on Windows, macOS, and Linux. Use for .xlsm/.xlsb/.xltm development, incremental VBA changes, portability review, and regression testing. Preserve the dedicated Windows Excel worker for automated editing, macro execution, and atomic release. Do not use for plain .xlsx data work or untrusted macro execution.
---

# usable-xlsm

## Select the OS first

Ask the user once which OS will **use the workbook**: Windows / macOS / Linux.
Reuse an already stated choice. Separately detect the agent's execution host;
a Linux agent generating a Windows workbook must use Linux static commands.
Do not ask again for each revision.

```text
uv run --project skills/usable-xlsm usable-xlsm environment --target-os windows
uv run --project skills/usable-xlsm usable-xlsm doctor
```

Run from the repository root; use the actual installed skill directory instead
of `skills/usable-xlsm` when installed elsewhere. Commands shown on one line
work in PowerShell, bash and zsh.

| Execution host | Available work | Runtime verification |
| --- | --- | --- |
| Windows + desktop Excel | Static work + existing transactional update/test/run | Dedicated Excel worker |
| macOS | Extraction, source editing, static checks, planning, new XLSM candidates | Manual Mac Excel import/compile/test; Windows worker for automated release |
| Linux / no Excel | Same static work and candidate creation | Optional LibreOffice compatibility smoke checks; Windows worker for Excel verification |

`doctor` defaults to the existing Excel readiness check on Windows and a
successful static capability check on Mac/Linux. `doctor --mode static` is
also available on Windows without Excel. `doctor --mode excel` failing must
block Excel jobs only, not extraction, source edits, or candidate creation.
Changing `--target-os` never enables COM on a non-Windows host.
Reuse the environment result for later revisions; rerun doctor when the host
or dependencies change, and before an Excel release.

Load only the needed guide:

- [windows-workflow.md](references/windows-workflow.md): Windows update/test/run,
  finalizers, atomic promotion, backups, and failure handling; preserve this path.
- [portable-workflow.md](references/portable-workflow.md): Mac/Linux handoff,
  LibreOffice/UNO setup and portability rules from real failure cases.
- [creation.md](references/creation.md): New sheets/buttons from a seed project
  or the bundled experimental MS-OVBA compressor/CFB writer.
- [security.md](references/security.md): Trust and signature rules for a new source.
- [setup.md](references/setup.md), [production.md](references/production.md):
  Provisioning a dedicated Windows worker, CI/signing, and monitoring.
- [vba-notes.md](references/vba-notes.md): Parser, component types, test harness.

## Cheap revision loop

1. Extract once with `extract --workbook book.xlsm --output work/vba`.
   Keep the exported sources; do not regenerate the workbook for a small fix.
2. Run `plan --source work/vba --workbook book.xlsm` to see module names and
   test discovery without printing all source. Add `--diff` only when needed.
   For a development iteration, `update --skip-unchanged` skips Excel when no
   source changes or finalizer are requested; it explicitly reports that tests
   did not run. Omit this option for release verification.
3. Read/edit only the relevant files. Use `check --source work/vba/Module1.bas`
   for fast feedback; run a full-directory check before handoff/release.
4. For a folder containing only changed **existing** modules, use
   `plan --partial` and Windows `update --partial`. Omitted modules are preserved.
   For one new standard module, use `plan --add-only --source NewModule.bas`
   and Windows `update --add-only` with the same file. Existing component names
   block addition, and all existing sources are verified after saving.
   Use the complete export and explicit `--sync` for other additions/deletions;
   never combine `--add-only`, `--partial` or `--sync`.
5. Use `check --source work/vba --target macos` or `--target libreoffice` for
   advisory portability findings; `--strict-portability` fails on warnings.
6. Run focused tests during development if useful, then the full isolated
   Windows suite before promotion. Do not silently skip tests or weaken trust.

Static checks validate syntax, not VBA references, compilation or runtime
behavior. Report which OS/runtime was actually tested and which checks remain.
Never call a candidate fully verified because static checks or LibreOffice pass.
Do not repeatedly retry an unsupported backend or install pywin32 on Mac/Linux.

## Preserve safety

Treat workbooks as executable code. Static inspection does not execute macros.
Use an approved policy or an explicit trust attestation before opening/running
any input. Never weaken Trust Center, remove MOTW, or add Trusted Locations.
Never call Excel COM directly; use the disposable, serialized Windows worker.
Keep updates on staging copies with re-extraction, isolated tests, backup,
audit and atomic promotion. Preserve signature/UserForm restrictions and
owned-PID cleanup. Never edit `.frm` text without its binary designer.
Do not use a portable builder to patch an existing workbook or compiled cache.
