"""Static trust gate for macro-enabled Excel workbooks.

Nothing in this module launches Excel.  Production callers must authorize a
workbook here before a COM worker is allowed to open it.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import hashlib
import json
import math
import os
from pathlib import Path
import subprocess
import sys
import tomllib
from typing import Any, Callable
import zipfile

from oletools.olevba import VBA_Parser


SUPPORTED_SUFFIXES = frozenset({".xlsm", ".xlsb", ".xltm"})
DEFAULT_SCAN_TIMEOUT = 30.0


def _validated_scan_timeout(value: float) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value <= 0:
        raise ValueError("scan_timeout must be a finite positive number")
    return float(value)


class WorkbookSecurityError(RuntimeError):
    """Raised when a workbook does not satisfy the execution policy."""

    def __init__(self, report: "PreflightReport") -> None:
        self.report = report
        self.error_code = "security_preflight"
        reasons = "; ".join(report.blocking_reasons) or "workbook is not authorized"
        super().__init__(f"Security preflight blocked {report.operation}: {reasons}")


@dataclass(frozen=True)
class SecurityFinding:
    code: str
    severity: str
    summary: str
    detail: str | None = None

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "code": self.code,
            "severity": self.severity,
            "summary": self.summary,
        }
        if self.detail:
            payload["detail"] = self.detail
        return payload


@dataclass(frozen=True)
class TrustPolicy:
    trusted_roots: tuple[Path, ...] = ()
    trusted_sha256: frozenset[str] = frozenset()
    allow_motw: bool = False
    allow_xlm: bool = False
    allow_vba_stomping: bool = False
    allow_scan_failures: bool = False
    allow_signed_updates: bool = False
    scan_timeout_seconds: float = DEFAULT_SCAN_TIMEOUT

    @classmethod
    def load(cls, path: str | Path | None) -> "TrustPolicy":
        if path is None:
            return cls()
        policy_path = Path(path)
        with policy_path.open("rb") as handle:
            document = tomllib.load(handle)
        raw = document.get("security", {})
        base = policy_path.parent.resolve()
        roots: list[Path] = []
        for value in raw.get("trusted_roots", []):
            candidate = Path(value)
            roots.append((base / candidate).resolve() if not candidate.is_absolute() else candidate.resolve())
        return cls(
            trusted_roots=tuple(roots),
            trusted_sha256=frozenset(str(value).lower() for value in raw.get("trusted_sha256", [])),
            allow_motw=bool(raw.get("allow_motw", False)),
            allow_xlm=bool(raw.get("allow_xlm", False)),
            allow_vba_stomping=bool(raw.get("allow_vba_stomping", False)),
            allow_scan_failures=bool(raw.get("allow_scan_failures", False)),
            allow_signed_updates=bool(raw.get("allow_signed_updates", False)),
            scan_timeout_seconds=_validated_scan_timeout(raw.get("scan_timeout_seconds", DEFAULT_SCAN_TIMEOUT)),
        )

    def trusts(self, workbook: Path, digest: str) -> bool:
        if digest.lower() in self.trusted_sha256:
            return True
        resolved = workbook.resolve()
        return any(resolved == root or resolved.is_relative_to(root) for root in self.trusted_roots)


@dataclass
class PreflightReport:
    workbook: Path
    operation: str
    sha256: str
    size: int
    trusted: bool
    allowed: bool
    trust_source: str | None = None
    has_vba: bool = False
    has_xlm: bool = False
    vba_stomping: bool = False
    signed_vba: bool = False
    signed_package: bool = False
    findings: list[SecurityFinding] = field(default_factory=list)
    blocking_reasons: list[str] = field(default_factory=list)
    scan_status: str = "not_run"
    scan_phase: str | None = None
    scan_timeout_seconds: float = DEFAULT_SCAN_TIMEOUT

    def to_dict(self) -> dict[str, Any]:
        return {
            "workbook": self.workbook.name,
            "operation": self.operation,
            "sha256": self.sha256,
            "size": self.size,
            "trusted": self.trusted,
            "trust_source": self.trust_source,
            "allowed": self.allowed,
            "has_vba": self.has_vba,
            "has_xlm": self.has_xlm,
            "vba_stomping": self.vba_stomping,
            "signed_vba": self.signed_vba,
            "signed_package": self.signed_package,
            "blocking_reasons": list(self.blocking_reasons),
            "findings": [finding.to_dict() for finding in self.findings],
            "scan_status": self.scan_status,
            "scan_phase": self.scan_phase,
            "scan_timeout_seconds": self.scan_timeout_seconds,
        }


def sha256_file(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def mark_of_the_web(path: str | Path) -> str | None:
    stream = Path(f"{Path(path)}:Zone.Identifier")
    try:
        value = stream.read_text(encoding="utf-8", errors="replace").strip()
    except OSError:
        return None
    return value or None


def _signature_flags(workbook: Path) -> tuple[bool, bool]:
    """Detect embedded VBA and package signature artifacts without Excel."""
    try:
        with zipfile.ZipFile(workbook) as package:
            names = {name.lower() for name in package.namelist()}
    except (OSError, zipfile.BadZipFile):
        return False, False
    signed_vba = any(name.endswith("vbaprojectsignature.bin") for name in names)
    signed_package = any(name.startswith("_xmlsignatures/") for name in names)
    return signed_vba, signed_package


def _inspect_macros(
    workbook: Path, *, allow_scan_failures: bool = False,
    progress: Callable[[dict], None] | None = None,
) -> dict:
    """Run the original full macro inspection inside the disposable scanner."""
    findings: list[SecurityFinding] = []
    blockers: list[str] = []
    has_vba = has_xlm = stomping = False
    vba_checked = False
    phase = "startup"

    def snapshot() -> dict:
        return {"has_vba": has_vba, "has_xlm": has_xlm, "stomping": stomping,
                "vba_checked": vba_checked, "phase": phase,
                "findings": [finding.to_dict() for finding in findings], "blockers": list(blockers)}

    def stage(name: str) -> None:
        nonlocal phase
        phase = name
        if progress is not None:
            progress(snapshot())

    parser: VBA_Parser | None = None
    try:
        stage("parser_open")
        parser = VBA_Parser(str(workbook))
        stage("vba_detection")
        has_vba = bool(parser.detect_vba_macros())
        vba_checked = True
        stage("xlm_scan")
        try:
            has_xlm = bool(parser.detect_xlm_macros())
        except Exception as exc:
            findings.append(SecurityFinding("xlm_scan_failed", "critical", "XLM scan failed", str(exc)))
            if not allow_scan_failures:
                blockers.append("XLM scan could not be completed")
        stage("stomping_scan")
        try:
            stomping = bool(parser.detect_vba_stomping()) if has_vba else False
        except Exception as exc:
            findings.append(SecurityFinding("stomping_scan_failed", "critical", "VBA stomping scan failed", str(exc)))
            if not allow_scan_failures:
                blockers.append("VBA stomping scan could not be completed")
        stage("macro_analysis")
        try:
            for kind, keyword, description in parser.analyze_macros(show_decoded_strings=True, deobfuscate=True) or []:
                normalized = str(kind).lower().replace(" ", "_")
                severity = "warning" if kind in {"AutoExec", "Suspicious", "IOC"} else "info"
                findings.append(SecurityFinding(f"olevba_{normalized}", severity, f"{kind}: {keyword}", str(description)))
        except Exception as exc:
            findings.append(SecurityFinding("macro_scan_failed", "critical", "Macro analysis failed", str(exc)))
            if not allow_scan_failures:
                blockers.append("macro analysis could not be completed")
    except Exception as exc:
        findings.append(SecurityFinding("workbook_scan_failed", "critical", "Workbook scan failed", str(exc)))
        blockers.append("workbook could not be statically inspected")
    finally:
        if parser is not None:
            stage("parser_close")
            parser.close()
    return {**snapshot(), "scan_status": "complete"}


def _scan_workbook(workbook: Path, policy: TrustPolicy, timeout: float) -> dict:
    # ponytail: one process per scan; pool only if small-book latency requires it.
    # Use the real interpreter rather than a Windows venv launcher, which can
    # leave its Python child running when only the launcher is killed.
    bootstrap = (
        "import json,sys; request=json.load(sys.stdin); "
        "sys.path=request.pop('sys_path'); "
        "from usable_xlsm._security_worker import scan; scan(request)"
    )
    state = {"has_vba": False, "has_xlm": False, "stomping": False,
             "vba_checked": False, "phase": "startup", "findings": [], "blockers": []}
    output = b""
    status = "failed"
    try:
        process = subprocess.run(
            [getattr(sys, "_base_executable", sys.executable), "-X", "utf8", "-c", bootstrap],
            input=json.dumps({"sys_path": sys.path, "workbook": str(workbook.resolve()),
                              "allow_scan_failures": policy.allow_scan_failures}).encode("utf-8"),
            capture_output=True, timeout=timeout,
            env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        )
        output = process.stdout
        status = "complete" if process.returncode == 0 else "failed"
    except subprocess.TimeoutExpired as exc:
        # subprocess.run kills and reaps the real interpreter before raising.
        output = exc.stdout or b""
        status = "timeout"
    except OSError:
        pass

    finished = False
    try:
        for line in output.decode("utf-8").splitlines():
            event = json.loads(line)
            result = event["state"]
            expected = {"has_vba", "has_xlm", "stomping", "vba_checked", "phase", "findings", "blockers"}
            if event["type"] == "result":
                expected.add("scan_status")
            if set(result) != expected or not isinstance(result["phase"], str):
                raise ValueError("invalid scanner result fields")
            if any(type(result[key]) is not bool for key in ("has_vba", "has_xlm", "stomping", "vba_checked")):
                raise ValueError("invalid scanner result flags")
            if not isinstance(result["findings"], list) or not isinstance(result["blockers"], list):
                raise ValueError("invalid scanner result lists")
            for finding in result["findings"]:
                item = SecurityFinding(**finding)
                if any(not isinstance(value, str) for value in (item.code, item.severity, item.summary)):
                    raise ValueError("invalid scanner finding")
                if item.detail is not None and not isinstance(item.detail, str):
                    raise ValueError("invalid scanner finding detail")
            if any(not isinstance(reason, str) for reason in result["blockers"]):
                raise ValueError("invalid scanner blocker")
            if event["type"] == "result" and result["scan_status"] == "complete":
                finished = True
            elif event["type"] != "progress":
                raise ValueError("invalid scanner event")
            state.update(result)
    except (UnicodeError, ValueError, KeyError, TypeError):
        if status != "timeout":
            status = "failed"
    if status == "complete" and (not finished or not state["vba_checked"]):
        status = "failed"
    if status == "complete":
        return state
    code = "macro_scan_timeout" if status == "timeout" else "macro_scan_worker_failed"
    reason = (f"workbook macro scan exceeded {timeout:g} seconds during {state['phase']}"
              if status == "timeout" else f"workbook macro scan failed during {state['phase']}")
    # Unknown XLM/stomping state is always a blocker, even for a policy that
    # accepts individual detector failures. No partial result authorizes a job.
    return {**state, "scan_status": status,
            "findings": [*state["findings"], SecurityFinding(code, "critical", reason).to_dict()],
            "blockers": [*state["blockers"], reason]}


def preflight_workbook(
    workbook_path: str | Path,
    *,
    operation: str = "execute",
    trust_workbook: bool = False,
    policy_path: str | Path | None = None,
    allow_signature_removal: bool = False,
    scan_timeout: float | None = None,
) -> PreflightReport:
    """Inspect and authorize a workbook without starting Excel.

    ``trust_workbook`` is an explicit operator assertion for a single command.
    A policy can instead trust stable roots or exact hashes.  Neither mechanism
    silently overrides high-risk format findings.
    """
    workbook = Path(workbook_path)
    if not workbook.is_file():
        raise FileNotFoundError(workbook)
    if operation not in {"inspect", "edit", "execute"}:
        raise ValueError(f"Unknown preflight operation: {operation}")

    policy = TrustPolicy.load(policy_path)
    timeout = _validated_scan_timeout(policy.scan_timeout_seconds if scan_timeout is None else scan_timeout)
    digest = sha256_file(workbook)
    policy_trusted = policy.trusts(workbook, digest)
    trusted = bool(trust_workbook or policy_trusted)
    trust_source = "operator" if trust_workbook else ("policy" if policy_trusted else None)
    findings: list[SecurityFinding] = []
    blockers: list[str] = []

    if workbook.suffix.lower() not in SUPPORTED_SUFFIXES:
        blockers.append(f"unsupported workbook type {workbook.suffix or '<none>'}")

    motw = mark_of_the_web(workbook)
    if motw:
        findings.append(SecurityFinding("motw", "critical", "Mark of the Web is present", motw))
        if not policy.allow_motw:
            blockers.append("Mark of the Web is present")

    signed_vba, signed_package = _signature_flags(workbook)
    if signed_vba:
        findings.append(SecurityFinding("signed_vba", "info", "VBA project signature is present"))
    if signed_package:
        findings.append(SecurityFinding("signed_package", "info", "Office package signature is present"))
    if operation == "edit" and (signed_vba or signed_package):
        if not (allow_signature_removal or policy.allow_signed_updates):
            blockers.append("editing would invalidate an existing digital signature")

    scan = _scan_workbook(workbook, policy, timeout)
    has_vba, has_xlm, stomping = (scan[key] for key in ("has_vba", "has_xlm", "stomping"))
    findings.extend(SecurityFinding(**item) for item in scan["findings"])
    blockers.extend(scan["blockers"])

    if has_xlm:
        findings.append(SecurityFinding("xlm_macro", "critical", "Excel 4.0/XLM macro is present"))
        if not policy.allow_xlm:
            blockers.append("Excel 4.0/XLM macros are not allowed")
    if stomping:
        findings.append(
            SecurityFinding("vba_stomping", "critical", "VBA source and compiled p-code differ")
        )
        if not policy.allow_vba_stomping:
            blockers.append("VBA stomping was detected")
    if not has_vba and not scan["blockers"]:
        findings.append(SecurityFinding("no_vba", "warning", "No VBA macros were detected"))

    if operation in {"edit", "execute"} and not trusted:
        blockers.append("source is not trusted; pass --trust-workbook or use an approved policy")

    # Keep output stable when multiple detectors report the same root cause.
    blockers = list(dict.fromkeys(blockers))
    return PreflightReport(
        workbook=workbook,
        operation=operation,
        sha256=digest,
        size=workbook.stat().st_size,
        trusted=trusted,
        trust_source=trust_source,
        allowed=not blockers,
        has_vba=has_vba,
        has_xlm=has_xlm,
        vba_stomping=stomping,
        signed_vba=signed_vba,
        signed_package=signed_package,
        findings=findings,
        blocking_reasons=blockers,
        scan_status=scan["scan_status"],
        scan_phase=scan["phase"],
        scan_timeout_seconds=timeout,
    )


def require_authorized(report: PreflightReport) -> PreflightReport:
    if not report.allowed:
        raise WorkbookSecurityError(report)
    return report
