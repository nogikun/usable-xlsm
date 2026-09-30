"""Describe capabilities without opening an Office document."""

from enum import Enum
import importlib.util
import shutil
import sys


class TargetOS(str, Enum):
    windows = "windows"
    macos = "macos"
    linux = "linux"


def host_os() -> str:
    return {"win32": "windows", "darwin": "macos"}.get(sys.platform, "linux")


def capabilities(target_os: TargetOS | str | None = None) -> dict:
    target = TargetOS(target_os or host_os()).value
    return {
        "host_os": host_os(),
        "target_os": target,
        "static_commands": ["extract", "check", "preflight", "plan", "build-project", "create", "verify"],
        "excel_worker_available": sys.platform == "win32",
        "excel_worker_commands": ["update", "test", "run"] if sys.platform == "win32" else [],
        "target_validation": "windows-excel-worker" if target == "windows" else "manual-excel-macos" if target == "macos" else "optional-libreoffice",
        "libreoffice": shutil.which("soffice") or shutil.which("libreoffice"),
        "next_step": "Use doctor --mode excel before update/test/run." if sys.platform == "win32" else "Continue static work; hand off source changes to a Windows Excel worker for automated update/test/run.",
        "runtime_verified": False,
    }


def static_dependencies() -> dict[str, bool]:
    return {name: importlib.util.find_spec(name) is not None for name in ("antlr4_vba", "oletools", "xlsxwriter")}
