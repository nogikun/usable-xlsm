"""Watchdog around the Excel COM worker.

A VBA compile error makes Excel pop a modal dialog inside a process that has no
visible window. The COM call never returns, ``DisplayAlerts = False`` does not
suppress it, and there is no in-band way to answer it. The only reliable escape
is to run every COM operation in a subprocess, give it a deadline, and kill it
from outside.

Before killing, we read the text off the dialog itself, which turns an opaque
hang into the actual VBA message (for example "Compile error: Expected End
Sub"). That message is the single most useful thing for fixing the code, so it
is worth the extra Win32 work.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import time
from typing import Any
import uuid

from .locking import LockBusyError, excel_host_lock

DEFAULT_TIMEOUT = 60.0

# How often to look for a blocking dialog while the worker runs.
DIALOG_POLL_INTERVAL = 0.75


class ExcelJobError(RuntimeError):
    """Raised when the Excel worker fails or has to be killed."""

    def __init__(self, message: str, *, dialog: str | None = None, timed_out: bool = False) -> None:
        self.dialog = dialog
        self.timed_out = timed_out
        super().__init__(message)


@dataclass
class JobResult:
    """Outcome of one COM job."""

    ok: bool
    data: Any = None
    error: str | None = None
    dialog: str | None = None
    timed_out: bool = False
    duration: float = 0.0
    stderr: str = ""
    job_id: str | None = None
    excel_pid: int | None = None
    office_version: str | None = None
    error_code: str | None = None


def _dialog_text(pid: int) -> str | None:
    """Read the text of any modal dialog owned by the given process.

    VBA error dialogs are standard Win32 dialogs (class #32770) whose message
    lives in child static controls, so the text is readable from outside even
    though the owning thread is blocked.
    """
    try:
        import win32gui
        import win32process
    except ImportError:
        return None

    captured: list[str] = []

    def collect_child(hwnd: int, acc: list[str]) -> bool:
        try:
            if win32gui.GetClassName(hwnd) in {"Static", "Button"}:
                text = win32gui.GetWindowText(hwnd)
                if text and text not in acc:
                    acc.append(text)
        except Exception:
            pass
        return True

    def visit(hwnd: int, _: Any) -> bool:
        try:
            if win32gui.GetClassName(hwnd) != "#32770":
                return True
            if win32process.GetWindowThreadProcessId(hwnd)[1] != pid:
                return True
            parts: list[str] = []
            title = win32gui.GetWindowText(hwnd)
            if title:
                parts.append(title)
            win32gui.EnumChildWindows(hwnd, collect_child, parts)
            captured.append(" | ".join(parts))
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(visit, None)
    except Exception:
        return None

    # Drop the button captions; they are noise once the message is in hand.
    cleaned = [
        text
        for text in captured
        if text and not text.strip() in {"OK", "Cancel", "Help", "はい", "いいえ"}
    ]
    return "\n".join(cleaned) if cleaned else None


def _kill(pid: int, *, tree: bool = True) -> None:
    command = ["taskkill", "/PID", str(pid)]
    if tree:
        command.append("/T")
    command.append("/F")
    subprocess.run(command, capture_output=True, check=False)


def find_stray_excel(before: set[int]) -> set[int]:
    """List Excel PIDs that were not running before we started."""
    try:
        # tasklist emits text in the console codepage, which is not necessarily
        # the interpreter's. Decode leniently: only the digits matter here, and
        # a decode error must never take down a cleanup path.
        raw = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            check=False,
        ).stdout
        output = (raw or b"").decode("utf-8", errors="replace")
    except Exception:
        return set()
    pids: set[int] = set()
    for line in output.splitlines():
        parts = [part.strip('" ') for part in line.split('","')]
        if len(parts) >= 2 and parts[1].isdigit():
            pids.add(int(parts[1]))
    return pids - before


def excel_pids() -> set[int]:
    """PIDs of every Excel process currently running."""
    return find_stray_excel(set())


def _run_job_unlocked(
    action: str,
    workbook: str | Path,
    *,
    job_id: str,
    timeout: float = DEFAULT_TIMEOUT,
    require_clean_host: bool = True,
    **fields: Any,
) -> JobResult:
    """Execute one COM job in a subprocess, killing it if it blocks."""
    # Close mkstemp's descriptor at once: Windows refuses to unlink a file that
    # still has an open handle, and the worker needs to write to it anyway.
    handle, raw_path = tempfile.mkstemp(prefix="usable-xlsm-", suffix=".worker.json")
    os.close(handle)
    pid_file = Path(raw_path)

    pre_existing = excel_pids()
    if require_clean_host and pre_existing:
        pid_file.unlink(missing_ok=True)
        return JobResult(
            ok=False,
            error=(
                "Production isolation check failed: Excel is already running "
                f"(PIDs {sorted(pre_existing)}). Use a dedicated Windows worker "
                "and close existing Excel instances before retrying."
            ),
            error_code="excel_host_not_clean",
            job_id=job_id,
        )

    job = {
        "action": action,
        "workbook": str(workbook),
        "pid_file": str(pid_file),
        "job_id": job_id,
        **fields,
    }

    process = subprocess.Popen(
        [sys.executable, "-m", "usable_xlsm._worker"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    def discard_pid_file() -> None:
        try:
            pid_file.unlink(missing_ok=True)
        except OSError:
            # A killed Excel can briefly keep the handle; leaving a stray file in
            # the temp directory is preferable to masking the real result.
            pass

    def read_excel_pid() -> int | None:
        try:
            payload = json.loads(pid_file.read_text(encoding="utf-8"))
            if payload.get("job_id") != job_id:
                return None
            raw = payload.get("excel_pid")
            return int(raw) if raw else None
        except Exception:
            return None

    # communicate() blocks, so it runs on its own thread while this one watches
    # for a dialog. Spotting the dialog ends the wait immediately instead of
    # burning the full timeout - in an edit/test loop that is the difference
    # between a three-second answer and a minute of dead air.
    transfer: dict[str, Any] = {}

    def pump() -> None:
        try:
            transfer["io"] = process.communicate(json.dumps(job))
        except Exception as exc:  # pragma: no cover - defensive
            transfer["exception"] = exc

    started = time.monotonic()
    thread = threading.Thread(target=pump, daemon=True)
    thread.start()

    dialog: str | None = None
    deadline = started + timeout
    while thread.is_alive() and time.monotonic() < deadline:
        thread.join(timeout=DIALOG_POLL_INTERVAL)
        if not thread.is_alive():
            break
        excel_pid = read_excel_pid()
        if excel_pid is None:
            continue
        found = _dialog_text(excel_pid)
        if found:
            # Any modal dialog here is fatal: nothing can click it, so the
            # worker is already stuck no matter what the dialog says.
            dialog = found
            break

    if thread.is_alive():
        duration = time.monotonic() - started
        excel_pid = read_excel_pid()
        ownership_failed = excel_pid is None
        if dialog is None and excel_pid is not None:
            dialog = _dialog_text(excel_pid)

        if excel_pid:
            _kill(excel_pid)
        else:
            # Never guess which Excel process belongs to us.  A broad sweep can
            # destroy a user's workbook or a concurrent job.  Dedicated worker
            # hosts are recycled if this rare ownership handshake fails.
            pass
        _kill(process.pid, tree=False)
        thread.join(timeout=5)
        discard_pid_file()

        if ownership_failed:
            message = (
                "Excel did not respond and its PID ownership handshake was not "
                "completed. No Excel process was guessed or killed; recycle the "
                "dedicated worker host before another production job."
            )
        elif dialog:
            message = (
                "Excel stopped on a modal dialog that no automation can dismiss, "
                f"so it was terminated after {duration:.1f}s.\nDialog text:\n{dialog}"
            )
        else:
            message = (
                f"Excel did not respond within {timeout:g}s and was terminated. "
                "This usually means a modal VBA dialog (a compile or runtime error)."
            )
        return JobResult(
            ok=False,
            error=message,
            dialog=dialog,
            timed_out=True,
            duration=duration,
            job_id=job_id,
            excel_pid=excel_pid,
            error_code=(
                "excel_pid_unknown"
                if ownership_failed
                else ("excel_modal_dialog" if dialog else "excel_timeout")
            ),
        )

    discard_pid_file()
    duration = time.monotonic() - started

    if "exception" in transfer:
        return JobResult(
            ok=False, error=f"Could not talk to the Excel worker: {transfer['exception']}",
            duration=duration, job_id=job_id, error_code="worker_transport_error",
        )

    stdout, stderr = transfer.get("io", ("", ""))

    payload: dict[str, Any] | None = None
    for line in reversed(stdout.strip().splitlines()):
        try:
            payload = json.loads(line)
            break
        except json.JSONDecodeError:
            continue

    if payload is None:
        return JobResult(
            ok=False,
            error=f"Excel worker produced no result.\nstdout: {stdout}\nstderr: {stderr}",
            duration=duration,
            stderr=stderr,
            job_id=job_id,
            error_code="worker_no_result",
        )

    if payload.get("job_id") != job_id:
        return JobResult(
            ok=False,
            error="Excel worker returned a result for the wrong job.",
            duration=duration,
            stderr=stderr,
            job_id=job_id,
            error_code="worker_job_mismatch",
        )

    if not payload.get("ok"):
        return JobResult(
            ok=False,
            error=payload.get("error") or "Unknown Excel worker failure",
            duration=duration,
            stderr=stderr,
            job_id=job_id,
            excel_pid=payload.get("excel_pid"),
            office_version=payload.get("office_version"),
            error_code=payload.get("error_code") or "excel_worker_error",
        )

    return JobResult(
        ok=True,
        data=payload.get("data"),
        duration=duration,
        stderr=stderr,
        job_id=job_id,
        excel_pid=payload.get("excel_pid"),
        office_version=payload.get("office_version"),
    )


def run_job(
    action: str,
    workbook: str | Path,
    *,
    timeout: float = DEFAULT_TIMEOUT,
    require_clean_host: bool = True,
    **fields: Any,
) -> JobResult:
    """Run one owned Excel job while holding the host-wide serialization lock."""
    job_id = str(fields.pop("job_id", uuid.uuid4().hex))
    try:
        with excel_host_lock(job_id):
            return _run_job_unlocked(
                action,
                workbook,
                job_id=job_id,
                timeout=timeout,
                require_clean_host=require_clean_host,
                **fields,
            )
    except LockBusyError as exc:
        return JobResult(
            ok=False,
            error=str(exc),
            error_code="excel_worker_busy",
            job_id=job_id,
        )
