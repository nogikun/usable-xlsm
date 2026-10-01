"""Small, read-only change reports for incremental VBA development."""

from difflib import unified_diff
from pathlib import Path
import re

from .core import _normalized_module_source, extract_vba
from .syntax import VBA_SUFFIXES
from .testing import discover


def read_sources(directory: str | Path) -> dict[str, str]:
    root = Path(directory)
    if not root.is_dir():
        raise NotADirectoryError(root)
    sources = {p.name: p.read_text(encoding="utf-8") for p in sorted(root.iterdir()) if p.is_file() and p.suffix.lower() in VBA_SUFFIXES}
    if not sources:
        raise ValueError(f"No .bas/.cls/.frm files found in {root}")
    if len({name.casefold() for name in sources}) != len(sources):
        raise ValueError("VBA component filenames must be unique ignoring case.")
    return sources


def plan_changes(source: str | Path, workbook: str | Path, *, partial: bool = False, include_diff: bool = False) -> dict:
    requested = read_sources(source)
    current = {Path(name).name: text for name, text in extract_vba(workbook).items()}
    changed = sorted(name for name in requested.keys() & current.keys() if _normalized_module_source(requested[name]) != _normalized_module_source(current[name]))
    added = sorted(requested.keys() - current.keys())
    removed = [] if partial else sorted(current.keys() - requested.keys())
    unchanged = sorted((requested.keys() & current.keys()) - set(changed))
    blocked = [name for name in changed + added if Path(name).suffix.lower() == ".frm"]
    if partial and added:
        blocked.extend(added)
    effective = {**current, **requested} if partial else requested
    discovery = discover(effective)
    result = {
        "ok": not blocked, "partial": partial,
        "changed": changed, "added": added, "removed": removed, "unchanged": unchanged,
        "blocked": sorted(set(blocked)), "requires_sync": bool(added or removed),
        "tests": [case.qualified for case in discovery.cases], "skipped_tests": discovery.skipped,
        "runtime_verified": False,
        "next_step": "Edit only the listed modules. Use update --partial for existing modules; review --sync explicitly for additions/deletions. Run the full test suite before release.",
    }
    if include_diff:
        result["diff"] = {name: "".join(unified_diff(_normalized_module_source(current.get(name, "")).splitlines(True), _normalized_module_source(requested.get(name, "")).splitlines(True), fromfile=f"before/{name}", tofile=f"after/{name}")) for name in changed + added + removed}
    return result


def _code_only(line: str) -> str:
    line = re.sub(r'"(?:[^"]|"")*"', lambda m: " " * len(m.group()), line)
    line = line.split("'", 1)[0]
    return re.split(r"(?:^|:)\s*Rem\b", line, maxsplit=1, flags=re.I)[0]


def portability_issues(sources: dict[str, str], target: str) -> list[dict]:
    """Advisory checks, deliberately separate from VBA syntax/compilation."""
    if target not in {"windows", "macos", "linux", "libreoffice"}:
        raise ValueError(f"Unknown portability target: {target}")
    rules = [("multi_statement", r":(?![=])", "Use one VBA statement per line; put Case and its body on separate lines.")]
    if target != "windows":
        rules += [
            ("error_constant", r"\bConst\b.*\bvbObjectError\b", "Use a numeric error literal; retain vbObjectError + n in a comment."),
            ("conditional_format", r"\bFormatConditions\b", "Provide an explicit fallback for this Excel feature."),
            ("activex", r"\b(?:CreateObject|GetObject|OLEObjects)\b", "Review COM/ActiveX use and provide a portable fallback."),
            ("windows_api", r"\bDeclare\b.*\bLib\b", "Guard native API declarations with platform conditionals and provide a fallback."),
        ]
    issues = []
    for name, text in sorted(sources.items()):
        for number, line in enumerate(text.splitlines(), 1):
            code = _code_only(line)
            code = re.sub(r"^\s*\w+:\s*$", "", code)
            for code_id, pattern, message in rules:
                if re.search(pattern, code, re.I):
                    issues.append({"file": name, "line": number, "code": code_id, "message": message, "severity": "warning"})
    return issues
