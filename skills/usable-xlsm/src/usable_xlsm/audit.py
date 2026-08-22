"""Privacy-conscious JSONL audit events for production workbook operations."""

from __future__ import annotations

from datetime import datetime, timezone
import json
from pathlib import Path
from typing import Any


AUDIT_DIR_NAME = ".usable-xlsm-audit"


def audit_path_for(workbook: str | Path) -> Path:
    path = Path(workbook)
    return path.parent / AUDIT_DIR_NAME / "audit.jsonl"


def write_audit_event(
    workbook: str | Path,
    *,
    job_id: str,
    action: str,
    status: str,
    sha256_before: str | None = None,
    sha256_after: str | None = None,
    duration: float | None = None,
    error_code: str | None = None,
    details: dict[str, Any] | None = None,
    audit_log: str | Path | None = None,
) -> Path:
    """Append an event without recording source code, cell values, or full paths."""
    book = Path(workbook)
    destination = Path(audit_log) if audit_log else audit_path_for(book)
    destination.parent.mkdir(parents=True, exist_ok=True)
    payload: dict[str, Any] = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "job_id": job_id,
        "action": action,
        "status": status,
        "workbook": book.name,
    }
    if sha256_before:
        payload["sha256_before"] = sha256_before
    if sha256_after:
        payload["sha256_after"] = sha256_after
    if duration is not None:
        payload["duration"] = round(duration, 3)
    if error_code:
        payload["error_code"] = error_code
    if details:
        payload["details"] = details
    with destination.open("a", encoding="utf-8", newline="\n") as handle:
        handle.write(json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str) + "\n")
    return destination
