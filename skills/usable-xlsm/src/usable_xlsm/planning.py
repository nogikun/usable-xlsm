"""Small, read-only change reports for incremental VBA development."""

from difflib import unified_diff
from pathlib import Path
import re

from .core import _normalized_module_source, extract_vba, module_changes, validate_vba_modules
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


def plan_changes(source: str | Path, workbook: str | Path, *, partial: bool = False, add_only: bool = False, include_diff: bool = False) -> dict:
    current = {Path(name).name: text for name, text in extract_vba(workbook).items()}
    requested = validate_vba_modules(source, workbook, partial=partial, add_only=True,
        current_sources=current) if add_only else read_sources(source)
    changes = module_changes(current, requested, partial=partial or add_only)
    changed, added, removed = (changes[key] for key in ("changed", "added", "removed"))
    blocked = [name for name in changed + added if Path(name).suffix.lower() == ".frm"]
    if partial and added:
        blocked.extend(added)
    effective = {**current, **requested} if partial or add_only else requested
    discovery = discover(effective)
    result = {
        "ok": not blocked, "partial": partial, "add_only": add_only, **changes,
        "blocked": sorted(set(blocked)), "requires_sync": bool(added or removed) and not add_only,
        "tests": [case.qualified for case in discovery.cases], "skipped_tests": discovery.skipped,
        "runtime_verified": False,
        "next_step": "Use update --add-only with this .bas file; existing modules will be preserved." if add_only else "Edit only the listed modules. Use update --partial for existing modules; review --sync explicitly for additions/deletions. Run the full test suite before release.",
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
