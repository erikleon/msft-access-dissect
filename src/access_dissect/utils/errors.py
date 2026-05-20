"""
Error hierarchy for access-dissect.

Design principle: extraction errors are isolated per-object.
One bad form/module/query must not abort the entire pipeline.
Only DatabaseOpenError and BitnessError are fatal (raised before extraction begins).
"""

from __future__ import annotations


class AccessDissectError(Exception):
    """Base class for all access-dissect errors."""


# ---------------------------------------------------------------------------
# Fatal errors — raised before extraction begins
# ---------------------------------------------------------------------------


class DatabaseOpenError(AccessDissectError):
    """Cannot open the Access database file.

    Causes: file not found, wrong password, file already open exclusively,
    corrupt file, or Access/ACE not installed.
    """


class BitnessError(AccessDissectError):
    """Python interpreter bitness does not match the installed ACE driver.

    The most common failure mode: 64-bit Python with 32-bit ACE, or vice versa.

    The error message includes remediation steps.
    """


class AccessNotInstalledError(AccessDissectError):
    """Microsoft Access (or ACE redistributable) is not installed or not registered in COM."""


# ---------------------------------------------------------------------------
# Per-object extraction errors — caught and recorded as warnings
# ---------------------------------------------------------------------------


class ExtractionError(AccessDissectError):
    """Base class for per-object extraction failures."""

    def __init__(self, object_type: str, object_name: str, message: str) -> None:
        self.object_type = object_type
        self.object_name = object_name
        super().__init__(f"[{object_type}:{object_name}] {message}")


class TableExtractionError(ExtractionError):
    def __init__(self, name: str, message: str) -> None:
        super().__init__("table", name, message)


class QueryExtractionError(ExtractionError):
    def __init__(self, name: str, message: str) -> None:
        super().__init__("query", name, message)


class FormExtractionError(ExtractionError):
    def __init__(self, name: str, message: str) -> None:
        super().__init__("form", name, message)


class ReportExtractionError(ExtractionError):
    def __init__(self, name: str, message: str) -> None:
        super().__init__("report", name, message)


class MacroExtractionError(ExtractionError):
    def __init__(self, name: str, message: str) -> None:
        super().__init__("macro", name, message)


class ModuleExtractionError(ExtractionError):
    def __init__(self, name: str, message: str) -> None:
        super().__init__("module", name, message)


# ---------------------------------------------------------------------------
# Specific condition errors — caught and recorded with a specific warning_code
# ---------------------------------------------------------------------------


class VBALockedError(AccessDissectError):
    """The VBA project is password-protected.

    HRESULT 0x800A03EC is the typical COM error code.
    Raised by ModuleExtractor; recorded as warning_code='VBA_LOCKED'.
    """


class LinkedTableError(AccessDissectError):
    """Attempted to traverse data from a linked (external) table.

    Linked tables are cataloged by metadata only; their data is never extracted.
    """


class WorkgroupSecurityError(AccessDissectError):
    """The database uses workgroup-level security (.mdw) and cannot be fully opened.

    Use --workgroup-file to supply the .mdw path.
    """


class COMSessionError(AccessDissectError):
    """The Access COM session is in a bad state and cannot be recovered."""


# ---------------------------------------------------------------------------
# Render / analysis errors
# ---------------------------------------------------------------------------


class RenderError(AccessDissectError):
    """Error during catalog rendering (Markdown, HTML, SQL generation)."""


class AnalysisError(AccessDissectError):
    """Error during catalog analysis (dependency graph, complexity scoring)."""


# ---------------------------------------------------------------------------
# Helper: format a COM error HRESULT for display
# ---------------------------------------------------------------------------


def format_hresult(hresult: int) -> str:
    """Format a COM HRESULT as a hex string."""
    return f"HRESULT {hresult & 0xFFFFFFFF:#010x}"
