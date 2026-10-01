# Repeated small edits

Create the first workbook as a candidate and validate it on its target runtime.
Then keep the workbook and exported VBA sources together through revisions:

```text
user changes one requirement
  -> edit only the affected source files
  -> check those files and plan --partial --diff
  -> update --partial: preflight -> staging -> verify -> all isolated tests
  -> repeat the same source with --skip-unchanged: preflight -> hash/lock/audit
  -> one new .bas module: plan/update --add-only
```

An unchanged iteration reports `runtime_verified: false` and `tests: null`.
It never counts as fresh runtime verification. Omit `--skip-unchanged` for the
final release. Use a trusted staging finalizer for worksheet/control changes;
do not modify the original workbook directly. Do not rebuild an existing
workbook with the portable candidate generator.

## Reproduce the benchmark

From a Git checkout with the baseline commit available:

```text
uv run --project skills/usable-xlsm python skills/usable-xlsm/tests/benchmark_incremental.py --repeats 2
uv run --project skills/usable-xlsm python skills/usable-xlsm/tests/benchmark_incremental.py --excel --repeats 2
```

The first command works on Windows/Linux/macOS and compares the full security
preflight and unchanged-update API path. The second requires provisioned
Windows Excel and runs three sequential edits: numeric function result,
caption function text, and a new standard module. Each edit runs all isolated
VBA tests; each identical repeat verifies zero Excel jobs and an unchanged
workbook hash. Each omitted module's source is compared after every update.
The script creates its own trusted fixtures and accepts no user workbooks.

The baseline is `eabda8e33ffaf1616d10fb6de1bad01ac879b5ee`; `--baseline` may
select another reviewed local revision. Repetitions alternate old/new order.
Reported times include security-worker and Excel-process startup, staging,
verification and tests. They exclude the benchmark interpreter startup and
fixture creation. Scores use `100 * before / after`; they compare the same
operation, not different operations or operating systems.

## Diagnosis and scope

The long benchmark was repeated expensive analysis, not a proven infinite
loop: the original full scan completed on a 1000-function input. Most time
was spent in olevba's string-expression decoder attempting to parse ordinary
numeric-code lines. Increasing or disabling packrat caching did not resolve
the 4-second reproducer. Candidate-line selection retains the same decoder
and complete-source detectors; see [security.md](security.md).

The 30-second macro-scan deadline prevents indefinite parser work from
authorizing an update. It does not bound filesystem operations, Excel jobs or
the entire multi-test workflow; those retain their existing controls.
Static cross-platform tests do not establish Mac Excel editing compatibility.

## Measurements (2026-10-02)

Windows 11 / Core i7-11700K / Python 3.13.5 / Excel 16.0 and Ubuntu WSL2 /
Python 3.13.5 used the same fixture: 22 modules, 12,857 UTF-8 source bytes.
Windows used two paired rounds; WSL used three for static/no-op operations.
The Windows workflow ran three edits and three repeated requests per round,
with seven passing isolated VBA tests per workflow (28 total across versions).
All old/new security reports matched exactly.

| Operation | Before (s) | After (s) | Time reduction | After score |
| --- | ---: | ---: | ---: | ---: |
| Windows: three edits + three unchanged repeats | 62.051 | 48.856 | 21.3% | 127.0 |
| Windows: full preflight, 22 modules | 2.561 | 0.824 | 67.8% | 310.6 |
| Windows: unchanged update, 22 modules | 2.713 | 0.673 | 75.2% | 403.0 |
| WSL: full preflight, 22 modules | 2.205 | 0.564 | 74.4% | 390.8 |
| WSL: unchanged update, 22 modules | 2.377 | 0.549 | 76.9% | 433.0 |
| Windows: unchanged update, 4 modules | 0.187 | 0.586 | -213.6% | 31.9 |
| WSL: unchanged update, 4 modules | 0.149 | 0.470 | -214.7% | 31.8 |

The original 52-module/1000-function fixture (69,344 source bytes) was also
recreated exactly for one bounded paired check: full preflight completed in
13.405 seconds before and 1.278 seconds after (10.49x), with identical reports.
This confirms that the previous multi-minute benchmark accumulated many slow
scans. It does not establish that all possible parser inputs terminate without
the new deadline.

The four-module fixture contains 851 source bytes (`--helpers 0`, three paired
rounds on each OS). Small inputs can be slower because every scan now pays for
a fresh process that can be terminated safely. The deadline is retained; this
is not a universal speedup claim. No persistent scan cache, process pool or
weaker inspection mode was introduced. The small sample counts support a
workload comparison, not a statistical guarantee for all user workbooks.

Windows and WSL passed 44 unit/regression tests each. WSL LibreOffice 26.2.5.2
also passed the separate generated-fixture smoke check from PR #7 after these
fixtures passed the new preflight: Japanese output, formula/format/button
behavior, positive and deliberate negative tests, save/reopen/re-execution,
and normalized VBA-source preservation. Raw VBA binary bytes changed during
LibreOffice save; arbitrary workbooks or signatures are not covered.
