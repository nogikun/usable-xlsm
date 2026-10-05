# Workbook trust and security policy

## Trust model

`preflight` runs before Excel starts. It calculates SHA-256 and checks MOTW,
embedded signature artifacts, VBA/XLM presence, AutoExec procedures,
suspicious APIs and indicators, obfuscation, and source/p-code mismatch.

Static extraction is permitted for review. Editing or execution additionally
requires either:

- an exact hash or containing root approved in a TOML policy; or
- `--trust-workbook`, representing an explicit operator attestation for that
  invocation.

Absence of MOTW is not evidence of trust.

Macro inspection runs in a disposable Python process with a 30-second deadline.
Timeouts, crashes and malformed scan results block the operation, including
unchanged updates, even if `allow_scan_failures` accepts individual detector
errors. No incomplete scan authorizes an Excel job. Hash, MOTW, signature and
trust-policy checks still run on every invocation; scan results are not cached.

For olevba 0.60.2, string-expression deobfuscation uses only complete logical
lines containing a double quote or `Chr` (case-insensitive). These cover every
leaf in that version's string-expression grammar. All other detectors receive
the full original source, including form/XLM/p-code strings. A different olevba
version uses its unfiltered decoder until its grammar is reviewed.

## Policy file

```toml
[security]
trusted_roots = ["D:/Approved/FinanceWorkbooks"]
trusted_sha256 = [
  "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef",
]

# Keep these false unless the organization's security owner has accepted the
# specific risk and the worker is disposable.
allow_motw = false
allow_xlm = false
allow_vba_stomping = false
allow_scan_failures = false
allow_signed_updates = false
scan_timeout_seconds = 30.0
```

Relative trusted roots resolve from the policy file's directory. Hash trust is
stronger for immutable releases; root trust is more convenient for a controlled
source repository or release share.

## Static scan diagnostics

The JSON report includes `scan_status`, `scan_phase`, and
`scan_timeout_seconds`. A timeout reports `macro_scan_timeout`, the last
observed phase (for example `macro_analysis` or `parser_close`), and
`allowed=false`; `preflight` exits with code 2. Crashes, startup failures, and
invalid or incomplete results report `macro_scan_worker_failed` and also deny
authorization. Partial detector flags are diagnostic evidence, not a completed
security assessment. Decoded string and deobfuscation checks remain enabled.

For a known large project, set `security.scan_timeout_seconds` in the policy
used by `update`, `test`, or `run`. Standalone `preflight` can override the policy
value with `--scan-timeout`, for example:

```powershell
uv run --project skills/usable-xlsm usable-xlsm preflight --workbook book.xlsm --operation inspect --scan-timeout 120
```

The scan budget must be finite and positive; it is independent of the Excel
worker's `--timeout`. The default remains 30 seconds. Neither
`--trust-workbook`, `allow_scan_failures=true`, nor skipping workbook tests
overrides an incomplete scan. Collect the blocked report and review the
source, or repeat the full scan with an explicitly larger budget.

## Hard blockers

Production defaults block:

- Mark of the Web;
- Excel 4.0/XLM macros;
- VBA stomping indicators;
- incomplete static scans;
- editing signed content without acknowledgement;
- an untrusted source for edit or execute operations.

AutoExec and suspicious keyword findings are retained in the report for review.
They are not independently fatal once the source itself is approved because
legitimate internal automation commonly uses events, file IO, and external
APIs. The policy owner decides whether such code belongs in an approved root or
hash list.

## Macro opening modes

Inspection and editing open with `Application.AutomationSecurity` set to
ForceDisable. Trusted execution uses an isolated worker and temporarily sets
the explicit execution mode only around `Workbooks.Open`, then restores the
previous value immediately. Events remain disabled unless explicitly enabled.

ForceDisable does not cover XLM macros, which is why XLM is rejected in static
preflight by default.

## Signatures

The tool detects package and VBA signature artifacts before editing. Any edit
invalidates the prior signature, so the safe default is to block. If a release
process deliberately uses `--allow-signature-removal`, treat the output as an
unsigned candidate and send it to the organization's controlled Office/VBA
code-signing service. Do not claim the resulting workbook is production-ready
until the final signature and signer policy have been independently verified.

Audit logs intentionally omit VBA source, argument values, cell values, and
full workbook paths. Protect the JSONL directory as operational evidence.
