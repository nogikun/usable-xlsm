"""Create a redacted diagnostic bundle without workbook contents."""

from __future__ import annotations

import json
import platform
from pathlib import Path
from typing import Any

from .audit import audit_path_for
from .core import MACRO_COPY_TAG, STAGING_COPY_TAG, TEST_COPY_TAG
from .runner import excel_pids
from ._version import __version__


def create_support_bundle(
    workbook_path: str | Path,
    destination: str | Path,
    *,
    audit_log: str | Path | None = None,
    max_events: int = 100,
) -> Path:
    workbook = Path(workbook_path)
    audit = Path(audit_log) if audit_log else audit_path_for(workbook)
    events: list[dict[str, Any]] = []
    if audit.is_file():
        lines = audit.read_text(encoding="utf-8", errors="replace").splitlines()
        selected = lines[-max_events:] if max_events else []
        for line in selected:
            try:
                item = json.loads(line)
            except json.JSONDecodeError:
                continue
            # Audit events are already redacted; enforce the boundary again.
            item.pop("path", None)
            item.pop("source", None)
            item.pop("cells", None)
            events.append(item)

    scratch = sorted(
        candidate.name
        for tag in (TEST_COPY_TAG, MACRO_COPY_TAG, STAGING_COPY_TAG)
        for candidate in workbook.parent.glob(f"*.{tag}.xls*")
    )
    payload = {
        "usable_xlsm_version": __version__,
        "platform": platform.platform(),
        "workbook": workbook.name,
        "excel_pids": sorted(excel_pids()),
        "scratch_files": scratch,
        "audit_events": events,
        "contains_workbook_data": False,
    }
    output = Path(destination)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    return output
