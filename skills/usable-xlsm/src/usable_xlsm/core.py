from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
import os
from pathlib import Path
import shutil
import time
from typing import Any
import uuid

from oletools.olevba import VBA_Parser

from .audit import write_audit_event
from .locking import workbook_lock
from .runner import DEFAULT_TIMEOUT, JobResult, excel_pids, run_job
from .security import (
    PreflightReport,
    WorkbookSecurityError,
    mark_of_the_web as _mark_of_the_web,
    preflight_workbook,
    require_authorized,
    sha256_file,
)
from .syntax import VBA_SUFFIXES, SyntaxIssue, VbaSyntaxError, check_directory
from .testing import (
    ASSERT_MODULE,
    ENTRY_POINT,
    RUNNER_MODULE,
    TestResult,
    assert_module_source,
    discover,
    generate_runner,
    parse_results,
)

# Suffixes used for scratch duplicates, shared with the sweeper. RESULT_COPY_TAG
# is deliberately not in that set: it marks a --save output the caller wants to
# keep, not a leftover to sweep away.
TEST_COPY_TAG = "tests"
MACRO_COPY_TAG = "testrun"
RESULT_COPY_TAG = "result"
STAGING_COPY_TAG = "staging"

# Modules the harness injects; nothing else should carry this prefix.
HARNESS_PREFIX = "UsableXlsm"

# Backups go in their own directory: they are worth keeping, but piling
# timestamped copies next to the workbook buries the file being worked on.
BACKUP_DIR_NAME = ".usable-xlsm-backups"


class VbaModuleMismatchError(ValueError):
    """Raised when source files do not match the workbook's VBA modules."""

    def __init__(self, missing: set[str], extra: set[str]) -> None:
        self.missing = missing
        self.extra = extra
        details = ["VBA module names do not match."]
        if missing:
            details.append(f"Missing from input directory: {', '.join(sorted(missing))}")
        if extra:
            details.append(f"Extra in input directory: {', '.join(sorted(extra))}")
        details.append("Pass sync=True (--sync) to add and remove modules instead.")
        super().__init__("\n".join(details))


class VbaUpdateError(RuntimeError):
    """Raised when Excel cannot update the VBA project."""

    def __init__(self, message: str, *, error_code: str = "vba_update_error") -> None:
        self.error_code = error_code
        super().__init__(message)


class VbaVerificationError(VbaUpdateError):
    """Raised when a staged workbook does not match the requested VBA source."""


def extract_vba(path: str | Path) -> dict[str, str]:
    """Extract VBA modules from an Excel macro-enabled workbook.

    Reads the compound file directly, so it needs neither Excel nor the VBA
    trust setting - safe to call at any point in the loop.
    """
    workbook = Path(path)
    if not workbook.is_file():
        raise FileNotFoundError(workbook)

    parser = VBA_Parser(str(workbook))
    try:
        return {
            vba_filename: code
            for _, _, vba_filename, code in parser.extract_macros()
        }
    finally:
        parser.close()


def validate_vba_modules(
    input_dir: str | Path,
    workbook_path: str | Path,
    *,
    sync: bool = False,
) -> dict[str, str]:
    """Read source files, checking their names against the workbook.

    Without ``sync`` the names must match exactly. That default exists because
    a typo in a filename would otherwise silently add a stray module or delete a
    real one, and this runs before Excel opens anything so a mismatch costs
    nothing to recover from.
    """
    source_dir = Path(input_dir)
    if not source_dir.is_dir():
        raise NotADirectoryError(source_dir)

    source_files = {
        path.name: path
        for path in source_dir.iterdir()
        if path.is_file() and path.suffix.lower() in VBA_SUFFIXES
    }
    if not source_files:
        raise ValueError(f"No .bas/.cls/.frm files found in {source_dir}")

    workbook_modules = extract_vba(workbook_path)
    if not sync:
        missing = set(workbook_modules) - set(source_files)
        extra = set(source_files) - set(workbook_modules)
        if missing or extra:
            raise VbaModuleMismatchError(missing, extra)

    return {
        name: path.read_text(encoding="utf-8")
        for name, path in sorted(source_files.items())
    }


def check_syntax(input_dir: str | Path) -> list[SyntaxIssue]:
    """Report syntax errors across a directory of VBA sources."""
    return check_directory(input_dir)


def backup_workbook(workbook_path: str | Path) -> Path:
    """Copy the workbook into a sibling backup directory, timestamped."""
    workbook = Path(workbook_path)
    directory = workbook.parent / BACKUP_DIR_NAME
    directory.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    destination = directory / (
        f"{workbook.stem}.{stamp}.{uuid.uuid4().hex[:8]}.bak{workbook.suffix}"
    )
    shutil.copy2(workbook, destination)
    return destination


def list_backups(workbook_path: str | Path) -> list[Path]:
    """Backups of one workbook, newest first."""
    workbook = Path(workbook_path)
    directory = workbook.parent / BACKUP_DIR_NAME
    if not directory.is_dir():
        return []
    return sorted(
        directory.glob(f"{workbook.stem}.*.bak{workbook.suffix}"),
        reverse=True,
    )


def prune_backups(workbook_path: str | Path, *, keep: int = 10) -> list[Path]:
    """Remove oldest managed backups, retaining the newest ``keep`` copies."""
    if keep < 1:
        raise ValueError("keep must be at least 1")
    removed: list[Path] = []
    for backup in list_backups(workbook_path)[keep:]:
        backup.unlink()
        removed.append(backup)
    return removed


def restore_workbook(
    workbook_path: str | Path,
    *,
    backup_path: str | Path | None = None,
    audit_log: str | Path | None = None,
) -> dict[str, Any]:
    """Atomically restore a managed backup without launching Excel."""
    workbook = Path(workbook_path)
    backup = Path(backup_path) if backup_path else next(iter(list_backups(workbook)), None)
    if backup is None or not backup.is_file():
        raise FileNotFoundError(backup or f"No backup found for {workbook}")
    job_id = uuid.uuid4().hex
    before = sha256_file(workbook) if workbook.is_file() else None
    staging = working_copy_path(workbook, STAGING_COPY_TAG, unique=True)
    destination = write_audit_event(
        workbook,
        job_id=job_id,
        action="restore",
        status="started",
        sha256_before=before,
        details={"backup": backup.name},
        audit_log=audit_log,
    )
    try:
        with workbook_lock(workbook, job_id):
            shutil.copy2(backup, staging)
            restored = sha256_file(staging)
            write_audit_event(
                workbook,
                job_id=job_id,
                action="restore",
                status="ready_to_restore",
                sha256_before=before,
                sha256_after=restored,
                details={"backup": backup.name},
                audit_log=destination,
            )
            os.replace(staging, workbook)
    finally:
        discard_working_copy(staging)
    audit_finalize_error: str | None = None
    try:
        write_audit_event(
            workbook,
            job_id=job_id,
            action="restore",
            status="succeeded",
            sha256_before=before,
            sha256_after=restored,
            details={"backup": backup.name},
            audit_log=destination,
        )
    except OSError as exc:
        audit_finalize_error = str(exc)
    return {
        "workbook": str(workbook),
        "backup": str(backup),
        "sha256": restored,
        "audit_log": str(destination),
        "audit_finalize_error": audit_finalize_error,
    }


def working_copy_path(workbook: Path, tag: str, *, unique: bool = False) -> Path:
    """Name for a scratch duplicate of a workbook.

    Deliberately beside the original rather than in the temp directory: macros
    routinely resolve paths from ``ThisWorkbook.Path``, and moving the copy
    elsewhere would quietly change their behaviour under test.

    ``unique`` inserts a short per-call token before the tag. Every ephemeral
    copy (run_tests, run_macro without --save) is deleted in that same call's
    ``finally``, so without a unique name two runs against the same workbook -
    concurrent test runs, or this repo's own eval harness running several
    agents at once - share one path: the second copy2 overwrites the first
    while Excel still has it open, and whichever run finishes first deletes the
    file out from under the other. The tag stays adjacent to the extension so
    the existing ``*.{tag}.xls*`` glob used by sweep_working_copies and
    verify_clean still matches.
    """
    if unique:
        token = f"{os.getpid():x}{uuid.uuid4().hex[:6]}"
        return workbook.with_name(f"{workbook.stem}.{token}.{tag}{workbook.suffix}")
    return workbook.with_name(f"{workbook.stem}.{tag}{workbook.suffix}")


def discard_working_copy(path: Path, attempts: int = 5, delay: float = 0.25) -> bool:
    """Delete a scratch copy, retrying briefly. Never raises.

    A force-killed Excel can hold the file open for a moment after the process
    is gone. Retrying handles that, and swallowing the final failure matters
    because this runs in a ``finally``: a locked file must not replace the
    result the caller actually asked for.
    """
    for attempt in range(attempts):
        try:
            path.unlink(missing_ok=True)
            return True
        except OSError:
            if attempt == attempts - 1:
                return False
            time.sleep(delay)
    return False


def _restore_backup(backup: Path, destination: Path, attempts: int = 5, delay: float = 0.25) -> bool:
    """Copy a backup over the workbook, retrying briefly. Never raises.

    Called right after a just-terminated Excel held the destination open, for
    the same reason ``discard_working_copy`` retries: a bare ``copy2`` here
    would let a transient ``PermissionError`` replace the compile/runtime error
    the caller is trying to report, which is the only thing worth surfacing.
    """
    for attempt in range(attempts):
        try:
            shutil.copy2(backup, destination)
            return True
        except OSError:
            if attempt == attempts - 1:
                return False
            time.sleep(delay)
    return False


def sweep_working_copies(directory: str | Path) -> list[Path]:
    """Delete scratch copies stranded by an interrupted run.

    Normal runs clean up after themselves; this exists for the case where the
    process itself was killed, which no ``finally`` can cover.
    """
    root = Path(directory)
    removed: list[Path] = []
    for tag in (TEST_COPY_TAG, MACRO_COPY_TAG, STAGING_COPY_TAG):
        for candidate in root.glob(f"*.{tag}.xls*"):
            if discard_working_copy(candidate):
                removed.append(candidate)
    return removed


@dataclass
class CleanlinessReport:
    """What the harness left behind, if anything."""

    workbook: Path | None
    harness_modules: list[str] = field(default_factory=list)
    scratch_copies: list[str] = field(default_factory=list)
    backups: list[str] = field(default_factory=list)
    excel_processes: list[int] = field(default_factory=list)
    unreadable: str | None = None

    @property
    def clean(self) -> bool:
        """Backups are deliberate, so they do not count against cleanliness."""
        return not (self.harness_modules or self.scratch_copies or self.unreadable)

    def problems(self) -> list[str]:
        issues: list[str] = []
        if self.unreadable:
            issues.append(f"could not read the workbook: {self.unreadable}")
        for name in self.harness_modules:
            issues.append(f"harness module left inside the workbook: {name}")
        for name in self.scratch_copies:
            issues.append(f"scratch copy left on disk: {name}")
        return issues


def verify_clean(
    target: str | Path,
    *,
    check_processes: bool = True,
) -> list[CleanlinessReport]:
    """Check that no test scaffolding survived a run.

    The harness injects modules into a workbook and duplicates it on disk. Both
    are supposed to disappear, but "supposed to" is not evidence - this reads the
    actual .xlsm and the actual directory so the claim can be checked rather
    than trusted. Accepts a single workbook or a directory of them.
    """
    root = Path(target)
    if root.is_dir():
        directory = root
        books = [
            path
            for path in sorted(root.glob("*.xls*"))
            if ".bak." not in path.name
            and f".{TEST_COPY_TAG}." not in path.name
            and f".{MACRO_COPY_TAG}." not in path.name
            and f".{STAGING_COPY_TAG}." not in path.name
        ]
    else:
        directory = root.parent
        books = [root]

    scratch = sorted(
        path.name
        for tag in (TEST_COPY_TAG, MACRO_COPY_TAG, STAGING_COPY_TAG)
        for path in directory.glob(f"*.{tag}.xls*")
    )
    backup_dir = directory / BACKUP_DIR_NAME
    backups = sorted(path.name for path in backup_dir.glob("*.bak.xls*")) if backup_dir.is_dir() else []
    processes = sorted(excel_pids()) if check_processes else []

    reports: list[CleanlinessReport] = []
    if not books:
        reports.append(
            CleanlinessReport(
                workbook=None,
                scratch_copies=scratch,
                backups=backups,
                excel_processes=processes,
            )
        )
        return reports

    for book in books:
        harness: list[str] = []
        unreadable: str | None = None
        try:
            harness = sorted(
                Path(name).stem
                for name in extract_vba(book)
                if Path(name).stem.startswith(HARNESS_PREFIX)
            )
        except Exception as exc:
            unreadable = str(exc)

        reports.append(
            CleanlinessReport(
                workbook=book,
                harness_modules=harness,
                # Directory-level facts repeat per workbook; that is fine, the
                # caller usually looks at one workbook at a time.
                scratch_copies=scratch,
                backups=backups,
                excel_processes=processes,
                unreadable=unreadable,
            )
        )
    return reports


def mark_of_the_web(workbook_path: str | Path) -> str | None:
    """Return the Zone.Identifier stream if Windows marked the file as foreign.

    Worth checking when a workbook "does not run its macros": a file that
    arrived by download, email or chat carries this mark, and Excel blocks its
    macros in the UI with "a potentially dangerous macro has been blocked".
    COM automation is unaffected, which is exactly what makes it confusing -
    these tools keep working while the user sees a blocked workbook.
    """
    return _mark_of_the_web(workbook_path)


def _normalized_module_source(source: str) -> str:
    lines = source.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    body = [line.rstrip() for line in lines if not line.startswith("Attribute ")]
    while body and not body[-1]:
        body.pop()
    return "\n".join(body)


def verify_applied_modules(
    workbook_path: str | Path,
    expected: dict[str, str],
    *,
    sync: bool,
    component_manifest: list[dict[str, Any]] | None = None,
) -> None:
    """Re-extract a closed staged workbook and compare the committed VBA."""
    actual = {Path(name).name: source for name, source in extract_vba(workbook_path).items()}
    expected_names = {Path(name).name for name in expected}
    if sync and component_manifest is not None:
        types_by_name = {str(item["name"]): int(item["type"]) for item in component_manifest}
        actual_editable = {
            f"{item['name']}{'.bas' if int(item['type']) == 1 else '.cls'}"
            for item in component_manifest
            if int(item["type"]) in {1, 2}
        }
        expected_editable = {
            Path(name).name
            for name in expected
            if Path(name).suffix.lower() in {".bas", ".cls"}
            and types_by_name.get(Path(name).stem, 1) in {1, 2}
        }
        missing = expected_editable - actual_editable
        extra = actual_editable - expected_editable
        if not missing and not extra:
            missing = set()
            extra = set()
        else:
            raise VbaVerificationError(
                "Staged editable-module manifest does not match: "
                f"missing={sorted(missing)}, extra={sorted(extra)}"
            )
    elif sync and not expected_names.issubset(set(actual)):
        missing = expected_names - set(actual)
        extra: set[str] = set()
    else:
        missing = set()
        extra = set()
    if missing or extra:
        raise VbaVerificationError(
            "Staged module manifest does not match: "
            f"missing={sorted(missing)}, extra={sorted(extra)}"
        )
    for name, source in expected.items():
        key = Path(name).name
        if key not in actual:
            raise VbaVerificationError(f"Staged workbook is missing module {key}")
        if _normalized_module_source(actual[key]) != _normalized_module_source(source):
            raise VbaVerificationError(f"Staged workbook did not preserve requested source for {key}")


def _require(result: JobResult, message: str) -> Any:
    if not result.ok:
        raise VbaUpdateError(
            f"{message}\n{result.error}",
            error_code=result.error_code or "excel_worker_error",
        )
    return result.data


def update_vba(
    input_dir: str | Path,
    workbook_path: str | Path,
    *,
    sync: bool = False,
    check: bool = True,
    backup: bool = True,
    timeout: float = DEFAULT_TIMEOUT,
    trust_workbook: bool = False,
    policy_path: str | Path | None = None,
    allow_signature_removal: bool = False,
    run_test_suite: bool = True,
    require_tests: bool = True,
    isolate_tests: bool = True,
    audit_log: str | Path | None = None,
    backup_keep: int = 10,
) -> dict[str, Any]:
    """Atomically promote a verified VBA update from a staging copy.

    The order here is deliberate: name validation and syntax checking both run
    before Excel is launched. Excel never opens the original for writing.
    """
    workbook = Path(workbook_path)
    if workbook.suffix.lower() not in {".xlsm", ".xlsb", ".xltm"}:
        raise ValueError(f"Output must be a macro-enabled workbook: {workbook}")

    preflight = require_authorized(
        preflight_workbook(
            workbook,
            operation="edit",
            trust_workbook=trust_workbook,
            policy_path=policy_path,
            allow_signature_removal=allow_signature_removal,
        )
    )
    requested_sources = validate_vba_modules(input_dir, workbook, sync=sync)
    current_sources = {Path(name).name: source for name, source in extract_vba(workbook).items()}
    for name, source in requested_sources.items():
        if Path(name).suffix.lower() != ".frm":
            continue
        key = Path(name).name
        if key not in current_sources:
            raise ValueError(f"Cannot create UserForm {key}; its paired .frx designer is required.")
        if _normalized_module_source(current_sources[key]) != _normalized_module_source(source):
            raise ValueError(
                f"Refusing to modify UserForm {key}; plain-text updates cannot preserve its designer."
            )
    module_sources = {
        name: source
        for name, source in requested_sources.items()
        if Path(name).suffix.lower() != ".frm"
    }

    if check:
        issues = check_directory(input_dir)
        if issues:
            raise VbaSyntaxError(issues)

    job_id = uuid.uuid4().hex
    started = time.monotonic()
    staging = working_copy_path(workbook, STAGING_COPY_TAG, unique=True)
    backup_path: Path | None = None
    audit_destination = write_audit_event(
        workbook,
        job_id=job_id,
        action="update",
        status="started",
        sha256_before=preflight.sha256,
        details={
            "trust_source": preflight.trust_source,
            "modules": sorted(Path(name).name for name in requested_sources),
            "sync": sync,
            "tests_required": bool(run_test_suite and require_tests),
        },
        audit_log=audit_log,
    )

    try:
        with workbook_lock(workbook, job_id):
            if sha256_file(workbook) != preflight.sha256:
                raise VbaUpdateError("Workbook changed after preflight; refusing a stale update.")

            backup_path = backup_workbook(workbook) if backup else None
            shutil.copy2(workbook, staging)
            result = run_job(
                "update",
                staging,
                job_id=job_id,
                timeout=timeout,
                modules=module_sources,
                sync=sync,
                security_mode="disable",
            )
            if not result.ok:
                raise VbaUpdateError(
                    result.error or "Staged update failed",
                    error_code=result.error_code or "excel_worker_error",
                )

            verify_applied_modules(
                staging,
                requested_sources,
                sync=sync,
                component_manifest=(result.data or {}).get("components"),
            )

            test_run: TestRun | None = None
            if run_test_suite:
                test_run = run_tests(
                    staging,
                    timeout=timeout,
                    trust_workbook=True,
                    require_tests=require_tests,
                    isolate=isolate_tests,
                    _skip_preflight=True,
                    _audit=False,
                )
                if not test_run.ok:
                    raise VbaUpdateError(
                        test_run.error
                        or f"Staged VBA tests failed: {len(test_run.failed)} failure(s)",
                        error_code=test_run.error_code or "tests_failed",
                    )

            if sha256_file(workbook) != preflight.sha256:
                raise VbaUpdateError("Workbook changed during staging; original was not replaced.")

            staged_hash = sha256_file(staging)
            write_audit_event(
                workbook,
                job_id=job_id,
                action="update",
                status="ready_to_promote",
                sha256_before=preflight.sha256,
                sha256_after=staged_hash,
                details={"tests": len(test_run.results) if test_run else None},
                audit_log=audit_destination,
            )
            os.replace(staging, workbook)

        duration = time.monotonic() - started
        removed_backups: list[Path] = []
        prune_error: str | None = None
        if backup:
            try:
                removed_backups = prune_backups(workbook, keep=backup_keep)
            except OSError as exc:
                # Promotion already succeeded. Retention housekeeping must not
                # turn a committed release into an ambiguous reported failure.
                prune_error = str(exc)
        audit_finalize_error: str | None = None
        try:
            write_audit_event(
                workbook,
                job_id=job_id,
                action="update",
                status="succeeded",
                sha256_before=preflight.sha256,
                sha256_after=staged_hash,
                duration=duration,
                details={
                    "backup": backup_path.name if backup_path else None,
                    "updated": (result.data or {}).get("updated", []),
                    "added": (result.data or {}).get("added", []),
                    "removed": (result.data or {}).get("removed", []),
                    "tests": len(test_run.results) if test_run else None,
                    "office_version": result.office_version,
                    "pruned_backups": len(removed_backups),
                    "backup_prune_error": prune_error,
                },
                audit_log=audit_destination,
            )
        except OSError as exc:
            # The mandatory ready_to_promote record was persisted before the
            # atomic replacement, so the release remains auditable.
            audit_finalize_error = str(exc)
        report = dict(result.data or {})
        report.update(
            {
                "backup": str(backup_path) if backup_path else None,
                "duration": duration,
                "job_id": job_id,
                "sha256_before": preflight.sha256,
                "sha256_after": staged_hash,
                "audit_log": str(audit_destination),
                "tests": len(test_run.results) if test_run else None,
                "office_version": result.office_version,
                "pruned_backups": len(removed_backups),
                "backup_prune_error": prune_error,
                "audit_finalize_error": audit_finalize_error,
            }
        )
        return report
    except Exception as exc:
        duration = time.monotonic() - started
        try:
            write_audit_event(
                workbook,
                job_id=job_id,
                action="update",
                status="failed",
                sha256_before=preflight.sha256,
                duration=duration,
                error_code=getattr(exc, "error_code", type(exc).__name__),
                details={"message": str(exc)[:1000]},
                audit_log=audit_destination,
            )
        except OSError:
            pass
        raise
    finally:
        discard_working_copy(staging)


def run_macro(
    workbook_path: str | Path,
    macro: str,
    *,
    args: list[Any] | None = None,
    read_cells: list[str] | None = None,
    save: bool = False,
    on_copy: bool = True,
    allow_in_place: bool = False,
    enable_events: bool = False,
    timeout: float = DEFAULT_TIMEOUT,
    trust_workbook: bool = False,
    policy_path: str | Path | None = None,
    audit_log: str | Path | None = None,
) -> JobResult:
    """Run a macro and collect its return value and any requested cell values.

    ``on_copy`` runs against a temporary duplicate so a macro with side effects
    can be exercised without dirtying the workbook under development. With
    ``save=True`` that duplicate is output the caller wants to keep, so it gets
    a name outside the scratch-copy pattern - otherwise verify_clean's cleanup
    and doctor.py --sweep would delete the very file the caller asked to keep,
    since both match on the same ``*.testrun.xls*`` pattern used for the
    ephemeral copies. ``save=False`` copies are always discarded, so they get a
    unique name instead: a fixed name would collide if this ran concurrently
    against the same workbook.
    """
    workbook = Path(workbook_path)
    if not on_copy and not allow_in_place:
        raise ValueError(
            "Production mode runs macros on a copy. Pass allow_in_place=True only "
            "for an explicitly approved interactive exception."
        )
    preflight = require_authorized(
        preflight_workbook(
            workbook,
            operation="execute",
            trust_workbook=trust_workbook,
            policy_path=policy_path,
        )
    )
    job_id = uuid.uuid4().hex
    started = time.monotonic()
    audit_destination = write_audit_event(
        workbook,
        job_id=job_id,
        action="run_macro",
        status="started",
        sha256_before=preflight.sha256,
        details={"macro": macro, "on_copy": on_copy, "save": save},
        audit_log=audit_log,
    )
    target = workbook
    temp: Path | None = None
    if on_copy:
        tag = RESULT_COPY_TAG if save else MACRO_COPY_TAG
        temp = working_copy_path(workbook, tag, unique=True)
        with workbook_lock(workbook, job_id):
            if sha256_file(workbook) != preflight.sha256:
                raise VbaUpdateError("Workbook changed after preflight; macro was not run.")
            shutil.copy2(workbook, temp)
        target = temp

    try:
        result = run_job(
            "run",
            target,
            job_id=job_id,
            timeout=timeout,
            macro=macro,
            args=args or [],
            read_cells=read_cells or [],
            save=save,
            enable_events=enable_events,
            security_mode="trusted_execute",
        )
        return result
    finally:
        cleanup_ok = True
        if temp is not None and not save:
            cleanup_ok = discard_working_copy(temp)
        if "result" in locals() and not cleanup_ok:
            result.ok = False
            result.error = f"Could not delete scratch copy {temp}"
            result.error_code = "cleanup_failed"
        final = locals().get("result")
        try:
            write_audit_event(
                workbook,
                job_id=job_id,
                action="run_macro",
                status="succeeded" if final is not None and final.ok else "failed",
                sha256_before=preflight.sha256,
                duration=time.monotonic() - started,
                error_code=None if final is not None and final.ok else (
                    final.error_code if final is not None else "run_aborted"
                ),
                details={
                    "macro": macro,
                    "saved_copy": temp.name if temp is not None and save else None,
                    "office_version": final.office_version if final is not None else None,
                    "excel_pid": final.excel_pid if final is not None else None,
                },
                audit_log=audit_destination,
            )
        except OSError as exc:
            if final is not None:
                final.ok = False
                final.error = f"Audit finalization failed: {exc}"
                final.error_code = "audit_failed"


@dataclass
class TestRun:
    """Outcome of a whole test run."""

    ok: bool
    results: list[TestResult] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)
    error: str | None = None
    dialog: str | None = None
    timed_out: bool = False
    duration: float = 0.0
    leftover: str | None = None
    error_code: str | None = None
    job_id: str | None = None
    audit_log: str | None = None
    office_versions: list[str] = field(default_factory=list)

    @property
    def passed(self) -> list[TestResult]:
        return [result for result in self.results if result.status == "pass"]

    @property
    def failed(self) -> list[TestResult]:
        return [result for result in self.results if result.status != "pass"]


def _run_test_discovery(
    workbook: Path,
    discovery: Any,
    *,
    timeout: float,
    job_id: str,
) -> TestRun:
    modules = {
        ASSERT_MODULE: assert_module_source(),
        RUNNER_MODULE: generate_runner(discovery),
    }

    # Always ephemeral, so a unique name avoids colliding with another test run
    # against the same workbook - e.g. two runs kicked off from a watcher, or
    # this repo's own eval harness running several agents in parallel.
    temp = working_copy_path(workbook, TEST_COPY_TAG, unique=True)
    shutil.copy2(workbook, temp)
    try:
        result = run_job(
            "run_tests",
            temp,
            job_id=f"{job_id}-{uuid.uuid4().hex[:8]}",
            timeout=timeout,
            modules=modules,
            replace=[RUNNER_MODULE],
            entry=ENTRY_POINT,
            security_mode="trusted_execute",
        )
    finally:
        leftover = None if discard_working_copy(temp) else str(temp)

    if not result.ok:
        return TestRun(
            ok=False,
            skipped=discovery.skipped,
            error=result.error,
            dialog=result.dialog,
            timed_out=result.timed_out,
            duration=result.duration,
            leftover=leftover,
            error_code=result.error_code,
            job_id=job_id,
            office_versions=[result.office_version] if result.office_version else [],
        )

    results = parse_results((result.data or {}).get("raw", ""))

    # A test that never reported back means the runner stopped early, which is
    # worth surfacing rather than quietly showing fewer results than expected.
    reported = {f"{item.module}.{item.name}" for item in results}
    for case in discovery.cases:
        if case.qualified not in reported:
            results.append(
                TestResult(
                    module=case.module,
                    name=case.name,
                    status="error",
                    duration=0.0,
                    message="test did not report a result",
                )
            )

    clean = leftover is None
    tests_ok = all(item.status == "pass" for item in results)
    return TestRun(
        ok=tests_ok and clean,
        results=results,
        skipped=discovery.skipped,
        duration=result.duration,
        leftover=leftover,
        error=(f"Could not delete scratch copy {leftover}" if leftover else None),
        error_code="cleanup_failed" if leftover else (None if tests_ok else "tests_failed"),
        job_id=job_id,
        office_versions=[result.office_version] if result.office_version else [],
    )


def run_tests(
    workbook_path: str | Path,
    *,
    pattern: str | None = None,
    timeout: float = DEFAULT_TIMEOUT,
    trust_workbook: bool = False,
    policy_path: str | Path | None = None,
    require_tests: bool = True,
    isolate: bool = True,
    audit_log: str | Path | None = None,
    _skip_preflight: bool = False,
    _audit: bool = True,
    _lock_workbook: bool = True,
) -> TestRun:
    """Run VBA tests on disposable copies, isolated per case by default."""
    workbook = Path(workbook_path)
    if _lock_workbook and not _skip_preflight:
        lock_job_id = uuid.uuid4().hex
        with workbook_lock(workbook, lock_job_id):
            return run_tests(
                workbook,
                pattern=pattern,
                timeout=timeout,
                trust_workbook=trust_workbook,
                policy_path=policy_path,
                require_tests=require_tests,
                isolate=isolate,
                audit_log=audit_log,
                _skip_preflight=False,
                _audit=_audit,
                _lock_workbook=False,
            )
    preflight: PreflightReport | None = None
    if not _skip_preflight:
        preflight = require_authorized(
            preflight_workbook(
                workbook,
                operation="execute",
                trust_workbook=trust_workbook,
                policy_path=policy_path,
            )
        )

    digest = preflight.sha256 if preflight else sha256_file(workbook)
    discovery = discover(extract_vba(workbook), pattern=pattern)
    job_id = uuid.uuid4().hex
    started = time.monotonic()
    audit_destination: Path | None = None
    if _audit:
        audit_destination = write_audit_event(
            workbook,
            job_id=job_id,
            action="run_tests",
            status="started",
            sha256_before=digest,
            details={"tests_discovered": len(discovery.cases), "isolated": isolate},
            audit_log=audit_log,
        )

    if not discovery.cases:
        run = TestRun(
            ok=not require_tests,
            results=[],
            skipped=discovery.skipped,
            error="No tests were discovered." if require_tests else None,
            error_code="no_tests" if require_tests else None,
            job_id=job_id,
        )
    elif isolate:
        combined: list[TestResult] = []
        office_versions: set[str] = set()
        leftovers: list[str] = []
        fatal_error: str | None = None
        fatal_code: str | None = None
        elapsed = 0.0
        for case in discovery.cases:
            single = type(discovery)(
                cases=[case],
                setups={case.module: discovery.setups.get(case.module, False)},
                teardowns={case.module: discovery.teardowns.get(case.module, False)},
                skipped=[],
            )
            one = _run_test_discovery(workbook, single, timeout=timeout, job_id=job_id)
            combined.extend(one.results)
            office_versions.update(one.office_versions)
            elapsed += one.duration
            if one.leftover:
                leftovers.append(one.leftover)
            if one.error and not one.results:
                fatal_error = one.error
                fatal_code = one.error_code
                break
        run = TestRun(
            ok=(
                fatal_error is None
                and not leftovers
                and len(combined) == len(discovery.cases)
                and all(item.status == "pass" for item in combined)
            ),
            results=combined,
            skipped=discovery.skipped,
            error=fatal_error or (f"Could not delete scratch copies: {leftovers}" if leftovers else None),
            duration=elapsed,
            leftover=", ".join(leftovers) if leftovers else None,
            error_code=(
                fatal_code
                or ("cleanup_failed" if leftovers else None)
                or ("tests_failed" if any(item.status != "pass" for item in combined) else None)
            ),
            job_id=job_id,
            office_versions=sorted(office_versions),
        )
    else:
        run = _run_test_discovery(workbook, discovery, timeout=timeout, job_id=job_id)

    run.audit_log = str(audit_destination) if audit_destination else None
    if _audit and audit_destination is not None:
        try:
            write_audit_event(
                workbook,
                job_id=job_id,
                action="run_tests",
                status="succeeded" if run.ok else "failed",
                sha256_before=digest,
                duration=time.monotonic() - started,
                error_code=run.error_code,
                details={
                    "tests": len(run.results),
                    "passed": len(run.passed),
                    "failed": len(run.failed),
                    "leftover": bool(run.leftover),
                    "office_versions": run.office_versions,
                },
                audit_log=audit_destination,
            )
        except OSError as exc:
            run.ok = False
            run.error = f"Audit finalization failed: {exc}"
            run.error_code = "audit_failed"
    return run


def inspect_workbook(
    workbook_path: str | Path,
    *,
    timeout: float = DEFAULT_TIMEOUT,
    trust_workbook: bool = False,
    policy_path: str | Path | None = None,
) -> dict[str, Any]:
    """List the workbook's VBA components and sheets via Excel."""
    report = require_authorized(
        preflight_workbook(
            workbook_path,
            operation="inspect",
            trust_workbook=trust_workbook,
            policy_path=policy_path,
        )
    )
    result = run_job(
        "inspect",
        workbook_path,
        timeout=timeout,
        security_mode="disable",
        read_only=True,
    )
    return _require(result, "Could not inspect the workbook.")
