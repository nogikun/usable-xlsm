"""Discover VBA tests, generate a runner for them, and read the results back.

The harness is convention-based: a ``Public Sub`` named ``Test_Something`` with
no arguments, in a standard module, is a test. Optional ``TestSetup`` and
``TestTeardown`` subs in the same module run around each of that module's tests.

Two things shape the generated runner. First, error handling lives in VBA, not
Python: an unhandled VBA error raises a modal dialog that would hang Excel, so
the runner wraps every call in ``On Error Resume Next`` and reports failures as
data. Second, the runner calls each test by name rather than dispatching through
``Application.Run``, because VBA compiles a procedure's static dependencies when
it is first called - naming the tests makes the run a compile check too.
"""

from __future__ import annotations

from dataclasses import dataclass
from importlib import resources
from pathlib import Path
import re

TEST_PREFIX = "Test_"
SETUP_NAME = "TestSetup"
TEARDOWN_NAME = "TestTeardown"

RUNNER_MODULE = "UsableXlsmRunner"
ASSERT_MODULE = "UsableXlsmAssert"
ENTRY_POINT = f"{RUNNER_MODULE}.UsableXlsmRunAll"

_PROCEDURE = re.compile(
    r"^[ \t]*(?:(?P<visibility>Public|Private|Friend)[ \t]+)?"
    r"(?:Static[ \t]+)?Sub[ \t]+(?P<name>\w+)[ \t]*\((?P<args>[^)]*)\)",
    re.IGNORECASE | re.MULTILINE,
)


@dataclass(frozen=True)
class TestCase:
    """One discovered test procedure."""

    module: str
    name: str

    @property
    def qualified(self) -> str:
        return f"{self.module}.{self.name}"


@dataclass(frozen=True)
class TestResult:
    """Outcome of one test procedure."""

    module: str
    name: str
    status: str  # pass | fail | error
    duration: float
    message: str

    @property
    def qualified(self) -> str:
        return f"{self.module}.{self.name}"

    @property
    def passed(self) -> bool:
        return self.status == "pass"


@dataclass(frozen=True)
class Discovery:
    """Everything found in a workbook's modules."""

    cases: list[TestCase]
    setups: dict[str, bool]
    teardowns: dict[str, bool]
    skipped: list[str]


def assert_module_source() -> str:
    """The bundled assertion module, as VBA source."""
    return resources.files("usable_xlsm").joinpath(
        "assets", f"{ASSERT_MODULE}.bas"
    ).read_text(encoding="utf-8")


def discover(modules: dict[str, str], pattern: str | None = None) -> Discovery:
    """Find test procedures across extracted modules.

    ``modules`` maps filenames (``Module1.bas``) to source, matching what
    ``extract_vba`` returns.
    """
    cases: list[TestCase] = []
    setups: dict[str, bool] = {}
    teardowns: dict[str, bool] = {}
    skipped: list[str] = []
    matcher = re.compile(pattern, re.IGNORECASE) if pattern else None

    for filename, source in sorted(modules.items()):
        module = Path(filename).stem
        is_standard = Path(filename).suffix.lower() == ".bas"

        for match in _PROCEDURE.finditer(source):
            name = match.group("name")
            visibility = (match.group("visibility") or "").lower()
            arguments = match.group("args").strip()

            if not name.lower().startswith(TEST_PREFIX.lower()):
                # A Private setup/teardown would still get called via a static
                # cross-module `Call Module.TestSetup` in the generated runner,
                # which is a compile error - so it must be excluded exactly like
                # a Private Test_* is excluded below, not silently registered.
                if visibility == "private" or arguments:
                    if name.lower() in (SETUP_NAME.lower(), TEARDOWN_NAME.lower()):
                        skipped.append(
                            f"{module}.{name} (Private or takes arguments; "
                            "TestSetup/TestTeardown must be Public with no arguments)"
                        )
                    continue
                if name.lower() == SETUP_NAME.lower():
                    setups[module] = True
                elif name.lower() == TEARDOWN_NAME.lower():
                    teardowns[module] = True
                continue

            # Each of these would fail at run time in a way that is much harder
            # to read than a message here, so report and move on.
            if not is_standard:
                skipped.append(f"{module}.{name} (only .bas modules are supported)")
                continue
            if visibility == "private":
                skipped.append(f"{module}.{name} (Private, cannot be called from the runner)")
                continue
            if arguments:
                skipped.append(f"{module}.{name} (takes arguments; tests must take none)")
                continue
            if matcher is not None and not matcher.search(f"{module}.{name}"):
                continue

            cases.append(TestCase(module, name))

    return Discovery(cases=cases, setups=setups, teardowns=teardowns, skipped=skipped)


def generate_runner(discovery: Discovery) -> str:
    """Build the VBA module that runs the discovered tests."""
    lines: list[str] = []
    for case in discovery.cases:
        lines.append("")
        lines.append("    Err.Clear")
        lines.append("    started = Timer")
        lines.append('    failurePhase = ""')
        if discovery.setups.get(case.module):
            lines.append(f"    Call {case.module}.{SETUP_NAME}")
            lines.append('    If Err.Number <> 0 Then failurePhase = "TestSetup"')
        # Skip the test if setup already failed, so the recorded message is the
        # setup's error rather than a confusing downstream one.
        lines.append(f"    If Err.Number = 0 Then Call {case.qualified}")
        lines.append('    If Err.Number <> 0 And Len(failurePhase) = 0 Then failurePhase = "test"')
        lines.append("    savedNumber = Err.Number")
        lines.append("    savedDescription = Err.Description")
        lines.append("    Err.Clear")
        if discovery.teardowns.get(case.module):
            lines.append(f"    Call {case.module}.{TEARDOWN_NAME}")
            lines.append("    If Err.Number <> 0 Then")
            lines.append("        If savedNumber = 0 Then")
            lines.append("            savedNumber = Err.Number")
            lines.append("            savedDescription = Err.Description")
            lines.append('            failurePhase = "TestTeardown"')
            lines.append("        Else")
            lines.append(
                '            savedDescription = savedDescription & " | TestTeardown: [" & '
                'CStr(Err.Number) & "] " & Err.Description'
            )
            lines.append("        End If")
            lines.append("    End If")
            lines.append("    Err.Clear")
        lines.append(
            f'    out = out & Record("{case.module}", "{case.name}", Timer - started, '
            "savedNumber, savedDescription, failurePhase)"
        )

    body = "\n".join(lines) if lines else "\n    ' no tests discovered"

    return f'''Attribute VB_Name = "{RUNNER_MODULE}"
Option Explicit

' Generated by usable-xlsm for a single test run, into a temporary copy of the
' workbook. Do not edit: it is rebuilt every time.
'
' The calls below are written out statically rather than dispatched by name so
' that VBA compiles every test as a dependency of this procedure.

Public Function UsableXlsmRunAll() As String
    Dim out As String
    Dim started As Single
    Dim savedNumber As Long
    Dim savedDescription As String
    Dim failurePhase As String

    On Error Resume Next
{body}

    On Error GoTo 0
    UsableXlsmRunAll = out
End Function

Private Function Record(ByVal moduleName As String, ByVal testName As String, _
                        ByVal elapsed As Single, ByVal errorNumber As Long, _
                        ByVal errorDescription As String, ByVal failurePhase As String) As String
    Dim status As String
    Dim message As String

    If errorNumber = 0 Then
        status = "pass"
    ElseIf errorNumber = {ASSERT_MODULE}.ASSERT_ERROR Then
        status = "fail"
        message = errorDescription
    Else
        status = "error"
        message = "[" & CStr(errorNumber) & "] " & errorDescription
    End If

    If Len(failurePhase) > 0 And failurePhase <> "test" Then
        message = failurePhase & ": " & message
    End If

    ' Tabs and newlines are the field and record separators, so they cannot
    ' survive inside a message.
    message = Replace(message, vbTab, " ")
    message = Replace(message, vbCr, " ")
    message = Replace(message, vbLf, " ")

    Record = moduleName & vbTab & testName & vbTab & status & vbTab & _
             Format$(elapsed, "0.000") & vbTab & message & vbLf
End Function
'''


def parse_results(raw: str) -> list[TestResult]:
    """Turn the runner's tab-separated output into structured results."""
    results: list[TestResult] = []
    for line in (raw or "").replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if not line.strip():
            continue
        parts = line.split("\t")
        if len(parts) < 5:
            continue
        module, name, status, elapsed, message = parts[0], parts[1], parts[2], parts[3], parts[4]
        try:
            duration = float(elapsed)
        except ValueError:
            duration = 0.0
        results.append(
            TestResult(
                module=module,
                name=name,
                status=status,
                duration=duration,
                message=message.strip(),
            )
        )
    return results
