"""Disposable static scanner; never starts Office or executes VBA."""
from contextlib import redirect_stdout
import json
from pathlib import Path
import sys

import oletools.olevba as olevba

from .security import _inspect_macros

_original_detect_strings = olevba.detect_vba_strings


def _candidate_vba_strings(code: str):
    # In olevba 0.60.2 every string-expression leaf is a quoted string or Chr.
    # Keep complete candidate lines, including comments/tabs/continued lines;
    # all other detectors still receive the complete original source.
    if olevba.__version__ != "0.60.2":
        return _original_detect_strings(code)
    candidates = "\n".join(line for line in code.splitlines() if '"' in line or 'chr' in line.lower())
    return _original_detect_strings(candidates)


def scan(request: dict) -> None:
    # This replacement lives only in this one-shot process, never in the caller.
    olevba.detect_vba_strings = _candidate_vba_strings
    with redirect_stdout(sys.stderr):
        result = _inspect_macros(Path(request["workbook"]), allow_scan_failures=request["allow_scan_failures"])
    print(json.dumps(result, ensure_ascii=False), flush=True)
