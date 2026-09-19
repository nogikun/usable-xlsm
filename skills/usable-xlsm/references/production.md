# Production operations

## Release transaction

An update follows this transaction:

```text
preflight -> source validation -> syntax check -> workbook lock -> backup
-> staging copy -> apply VBA -> optional trusted post-macro on staging -> close
-> static re-extraction, worksheet/cell/button package verification, and manifest/source comparison
-> isolated Test_* runs -> original hash recheck -> atomic replace -> audit
```

The original is never opened by Excel for writing. A failure before atomic
replacement leaves it unchanged. Failures during preflight or input validation
may occur before a backup is created; they still leave the original unchanged
and emit a failed audit event when an audit destination is available. A
`post_macro` is run only on the staging copy and is enabled explicitly by the caller; without one, the update worker
opens the staging workbook with macros disabled. Backups live under `.usable-xlsm-backups/`;
audit events live under `.usable-xlsm-audit/audit.jsonl` by default.

Before replacement, the tool persists a `ready_to_promote` event containing the
original and staged hashes. Final success logging is best effort only after that
mandatory commit record exists. The default backup retention is ten copies and
can be changed with `--backup-keep`.

`restore` copies a backup to a new staging file and atomically replaces the
target while holding the workbook lock.

## Development versus release

The release transaction is intentionally more expensive than an edit loop.
During development, apply to a disposable copy with `--no-test`, then run one
known smoke macro or one filtered `Test_*` case on that copy. The apply itself
still uses real Excel and performs save/close/package checks. Run the complete
isolated suite only for release or handoff; do not make every edit pay for
unrelated tests.

## Queue and recovery

Use an external queue with concurrency one per worker. The package also holds a
host-wide lock and a workbook-specific lock as defense in depth.

The watchdog handshake contains a job ID, worker PID, and exact Excel PID. On a
timeout it terminates only that Excel process tree. If the handshake is absent,
it stops the Python worker, reports `excel_pid_unknown`, and requires VM
recycle; it never sweeps newly observed Excel processes.

Recommended orchestrator response:

| Error | Response |
|---|---|
| `excel_worker_busy` | Leave queued; retry with bounded backoff. |
| `excel_host_not_clean` | Quarantine and recycle worker. |
| `excel_modal_dialog` | Fail release; report captured dialog text. |
| `excel_pid_unknown` | Quarantine and recycle worker. |
| `cleanup_failed` | Fail release; quarantine workspace and verify residue. |
| security exit `2` | Do not retry automatically; require review or policy change. |

## CI matrix

Test the supported combinations actually deployed:

- Office 32-bit and/or 64-bit;
- Japanese and English Office if both are supported;
- each Office update channel in the worker pool;
- signed, unsigned, MOTW, XLM, password-protected, locked-VBProject, UserForm,
  external-link, event-handler, compile-error, modal-dialog, and timeout cases.

Fault-injection checks must prove that the original hash is unchanged, the exit
code is non-zero, no scratch file survives, and only the owned Excel PID can be
terminated.

## Metrics and retention

Collect from JSONL events without logging workbook contents:

- success and security-block rate;
- duration percentiles;
- modal-dialog and timeout rate;
- unknown-PID and dirty-host rate;
- cleanup and restore rate;
- Office version per worker image.

Define retention separately for backups and audit logs. Backups contain the
full workbook and therefore inherit its data classification.

## Supported boundary

This design supports trusted internal workbooks on dedicated desktop Excel
workers. It is not a shared web-service architecture. For high-concurrency
server-side document processing, use Open XML, Microsoft Graph, or another
server-safe format/API and reserve this worker for the VBA-specific release and
execution step.
