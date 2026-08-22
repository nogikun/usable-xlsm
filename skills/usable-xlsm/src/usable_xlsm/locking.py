"""Small cross-process locks used to serialize Excel automation safely."""

from __future__ import annotations

from contextlib import AbstractContextManager
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import subprocess
import tempfile
from typing import Any


class LockBusyError(RuntimeError):
    """Raised when another usable-xlsm job owns a lock."""


def _pid_is_running(pid: int) -> bool:
    if pid <= 0:
        return False
    if os.name == "nt":
        try:
            output = subprocess.run(
                ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
                capture_output=True,
                check=False,
            ).stdout.decode(errors="replace")
            return f'"{pid}"' in output
        except OSError:
            return True
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


class FileLock(AbstractContextManager["FileLock"]):
    def __init__(self, path: str | Path, *, purpose: str, job_id: str) -> None:
        self.path = Path(path)
        self.purpose = purpose
        self.job_id = job_id
        self._owned = False

    def _payload(self) -> dict[str, Any]:
        return {
            "pid": os.getpid(),
            "job_id": self.job_id,
            "purpose": self.purpose,
            "created_at": datetime.now(timezone.utc).isoformat(),
        }

    def _remove_if_stale(self) -> bool:
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
            owner = int(payload.get("pid", 0))
        except Exception:
            return False
        if _pid_is_running(owner):
            return False
        try:
            self.path.unlink()
            return True
        except OSError:
            return False

    def acquire(self) -> "FileLock":
        self.path.parent.mkdir(parents=True, exist_ok=True)
        for attempt in range(2):
            try:
                descriptor = os.open(self.path, os.O_WRONLY | os.O_CREAT | os.O_EXCL)
            except FileExistsError:
                if attempt == 0 and self._remove_if_stale():
                    continue
                try:
                    owner = self.path.read_text(encoding="utf-8", errors="replace")
                except OSError:
                    owner = "unreadable lock"
                raise LockBusyError(f"Another usable-xlsm job holds {self.path}: {owner}")
            else:
                with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
                    json.dump(self._payload(), handle, ensure_ascii=False)
                self._owned = True
                return self
        raise LockBusyError(f"Could not acquire {self.path}")

    def release(self) -> None:
        if not self._owned:
            return
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
            if payload.get("job_id") == self.job_id:
                self.path.unlink(missing_ok=True)
        finally:
            self._owned = False

    def __enter__(self) -> "FileLock":
        return self.acquire()

    def __exit__(self, exc_type: object, exc: object, traceback: object) -> None:
        self.release()


def excel_host_lock(job_id: str) -> FileLock:
    return FileLock(
        Path(tempfile.gettempdir()) / "usable-xlsm" / "excel-worker.lock",
        purpose="exclusive Excel COM worker",
        job_id=job_id,
    )


def workbook_lock(workbook: str | Path, job_id: str) -> FileLock:
    path = Path(workbook)
    return FileLock(
        path.with_name(f".{path.name}.usable-xlsm.lock"),
        purpose="workbook transaction",
        job_id=job_id,
    )
