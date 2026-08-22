# Dedicated worker setup

Use a dedicated Windows user or disposable VM with a licensed desktop Excel
installation. Do not run this worker on a person's everyday Excel session.

## Required profile setting

Excel blocks programmatic VBA project writes unless this worker profile enables:

Excel > File > Options > Trust Center > Trust Center Settings > Macro Settings
> Trust access to the VBA project object model.

Japanese UI:

ファイル > オプション > トラスト センター > トラスト センターの設定 >
マクロの設定 > 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」

Registry evidence for Office 16:

```text
HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1
```

This setting is a security exposure. Keep it off normal user profiles and scope
it to the dedicated worker account. Do not change it automatically from a job.

## Installation and health check

```powershell
uv sync --project skills/usable-xlsm --frozen
uv run --project skills/usable-xlsm usable-xlsm doctor
```

The health check must report Windows, `pywin32`, `antlr4-vba`, enabled
`AccessVBOM`, and no running Excel process. It never kills Excel processes.

## Worker invariants

- One queued job at a time per Windows profile/VM.
- No interactive Excel work on that profile.
- Fixed Office architecture and update channel per worker pool.
- Recycle the worker after a timeout, unknown PID, COM crash, or health-check
  failure.
- Do not install unnecessary COM/VBE add-ins; they share Excel global state and
  can introduce dialogs or startup code.
- Service accounts need a real initialized user profile and writable Office
  profile directories. Do not host Excel in IIS, a Windows service, or another
  non-interactive server context.

Extraction, syntax checking, and preflight are static and can run without
Excel. Editing and execution require this Windows worker.
