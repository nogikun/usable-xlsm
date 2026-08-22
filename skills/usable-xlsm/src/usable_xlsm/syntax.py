"""Grammar-based syntax checks for VBA source, run before Excel opens the file.

Excel accepts whatever bytes we hand it without complaint; a compile error only
surfaces later as a modal dialog inside an invisible Excel process, which COM
cannot dismiss. Catching mistakes here is what keeps the edit/apply/test loop
from deadlocking.

Parsing uses the ANTLR4 VBA grammar from ``antlr4-vba`` rather than a
hand-rolled scanner, so constructs like single-line ``If``, line continuations
and ``Select Case`` are handled by a real grammar instead of heuristics.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re

from antlr4 import CommonTokenStream, InputStream
from antlr4.error.ErrorListener import ErrorListener

from antlr4_vba.vbaLexer import vbaLexer
from antlr4_vba.vbaParser import vbaParser


VBA_SUFFIXES = frozenset({".bas", ".cls", ".frm"})

# The grammar's start rule expects a well-formed module header. Files exported
# by olevba carry only bare Attribute lines, and some of those (VB_Base) are not
# in the grammar at all, so we rebuild a canonical header instead of trusting
# whatever the export produced.
_CLASS_HEADER = ("VERSION 1.0 CLASS", "BEGIN", "  MultiUse = -1  'True", "END")


@dataclass(frozen=True)
class SyntaxIssue:
    """One problem found in a VBA source file."""

    file: str
    line: int
    column: int
    message: str

    def __str__(self) -> str:
        return f"{self.file}:{self.line}:{self.column}: {self.message}"


class VbaSyntaxError(ValueError):
    """Raised when VBA source fails the static checks."""

    def __init__(self, issues: list[SyntaxIssue]) -> None:
        self.issues = issues
        body = "\n".join(f"  {issue}" for issue in issues)
        super().__init__(f"VBA syntax check failed:\n{body}")


class _Collector(ErrorListener):
    """Gathers parser errors instead of printing them to stderr."""

    def __init__(self) -> None:
        self.errors: list[tuple[int, int, str]] = []

    def syntaxError(self, recognizer, offendingSymbol, line, column, msg, e) -> None:  # noqa: N802
        self.errors.append((line, column, msg))


def _condense(message: str) -> str:
    """Shorten ANTLR's exhaustive 'expecting {...}' token dumps.

    The full list can run to hundreds of tokens, which buries the useful part of
    the message when it reaches a human or a model reading tool output.
    """
    match = re.search(r"expecting \{(.+)\}", message, re.DOTALL)
    if match is None:
        return message
    tokens = [token.strip() for token in match.group(1).split(",")]
    shown = ", ".join(tokens[:5])
    if len(tokens) > 5:
        shown += f", ... (+{len(tokens) - 5} more)"
    return message[: match.start()] + f"expecting one of: {shown}"


def _normalize(source: str, suffix: str, module_name: str) -> tuple[str, int]:
    """Rebuild a parseable module header, preserving original line numbers.

    Attribute lines are blanked rather than deleted so reported line numbers
    still match the file on disk, and because ``update_vba`` strips them before
    writing anyway - this checks the text that actually reaches Excel.
    """
    lines = source.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    body = ["" if line.startswith("Attribute ") else line for line in lines]
    header = list(_CLASS_HEADER) if suffix in {".cls", ".frm"} else []
    header.append(f'Attribute VB_Name = "{module_name}"')
    return "\r\n".join(header + body), len(header)


def check_source(
    source: str,
    filename: str = "<source>",
    module_name: str | None = None,
    suffix: str | None = None,
) -> list[SyntaxIssue]:
    """Check one VBA module for syntax errors."""
    resolved_suffix = (suffix or Path(filename).suffix or ".bas").lower()
    resolved_name = module_name or Path(filename).stem or "Module1"

    text, offset = _normalize(source, resolved_suffix, resolved_name)

    collector = _Collector()
    lexer = vbaLexer(InputStream(text))
    lexer.removeErrorListeners()
    lexer.addErrorListener(collector)
    parser = vbaParser(CommonTokenStream(lexer))
    parser.removeErrorListeners()
    parser.addErrorListener(collector)
    parser.startRule()

    return [
        SyntaxIssue(filename, max(line - offset, 1), column, _condense(message))
        for line, column, message in collector.errors
    ]


def check_file(path: str | Path) -> list[SyntaxIssue]:
    """Check a single exported VBA file."""
    target = Path(path)
    return check_source(
        target.read_text(encoding="utf-8"),
        filename=target.name,
        module_name=target.stem,
        suffix=target.suffix.lower(),
    )


def check_directory(path: str | Path) -> list[SyntaxIssue]:
    """Check every VBA source file in a directory."""
    directory = Path(path)
    if not directory.is_dir():
        raise NotADirectoryError(directory)
    issues: list[SyntaxIssue] = []
    for item in sorted(directory.iterdir()):
        if item.is_file() and item.suffix.lower() in VBA_SUFFIXES:
            issues.extend(check_file(item))
    return issues
