# usable-xlsm

Production-oriented tooling for reviewing, updating, and testing VBA in trusted
Excel macro-enabled workbooks.

Key properties:

- static security preflight before Excel starts;
- explicit trust policy or per-invocation attestation;
- one owned Excel process on a dedicated worker;
- staged updates with backup, re-extraction verification, isolated tests, and
  atomic promotion;
- JSON output, stable error codes, privacy-conscious JSONL audit events;
- no broad Excel process termination;
- MOTW, XLM, VBA stomping, signatures, UserForms, teardown failures, and cleanup
  failures handled explicitly.

Install and inspect the command surface:

```powershell
uv sync --project skills/usable-xlsm --frozen
uv run --project skills/usable-xlsm usable-xlsm --help
```

Start with:

```powershell
uv run --project skills/usable-xlsm usable-xlsm doctor
uv run --project skills/usable-xlsm usable-xlsm preflight `
  --workbook book.xlsm --operation edit --policy usable-xlsm.toml
```

Operational instructions live in `SKILL.md`. Trust policy, dedicated worker,
and release details are in `references/security.md`,
`references/setup.md`, and `references/production.md`.
