"""
ReportExtractor: extract ReportDef objects by opening each report in design view.

Follows the same pattern as FormExtractor: DoCmd.OpenReport in design view,
enumerate sections and controls, close. Same timeout and recovery mechanisms.
"""

from __future__ import annotations

import threading
from typing import Any

from access_dissect.catalog.constants import (
    REPORT_SECTION_NAMES,
    control_type_name,
)
from access_dissect.catalog.models import (
    ControlDef,
    ControlGeometry,
    ExtractionWarning,
    ReportDef,
    ReportGrouping,
    ReportSection,
    WarningSeverity,
)
from access_dissect.extract.forms import (
    _extract_control,
    _safe_bool,
    _safe_int,
    _safe_str,
)

# Access constants
AC_REPORT = 3
AC_DESIGN = 1


def _open_report_with_timeout(
    app: Any,
    report_name: str,
    timeout_seconds: float = 30.0,
) -> bool:
    """Open a report in design view with a timeout."""
    success = [False]
    exception_holder: list[Exception] = []

    def _open() -> None:
        try:
            app.DoCmd.OpenReport(report_name, AC_DESIGN)
            success[0] = True
        except Exception as e:
            exception_holder.append(e)

    t = threading.Thread(target=_open, daemon=True)
    t.start()
    t.join(timeout=timeout_seconds)

    if t.is_alive():
        return False

    if exception_holder:
        raise exception_holder[0]

    return success[0]


def _extract_groupings(report_obj: Any) -> list[ReportGrouping]:
    """Extract grouping/sorting levels from a report's GroupLevel collection."""
    groupings: list[ReportGrouping] = []
    try:
        # GroupLevel collection is 0-indexed; Count property tells us how many
        count = int(report_obj.GroupLevel.Count) if hasattr(report_obj, "GroupLevel") else 0
        for i in range(count):
            try:
                gl = report_obj.GroupLevel(i)
                groupings.append(
                    ReportGrouping(
                        field_or_expression=_safe_str(gl, "ControlSource") or "",
                        group_on=_safe_int(gl, "GroupOn"),
                        group_interval=_safe_int(gl, "GroupInterval"),
                        keep_together=_safe_bool(gl, "KeepTogether"),
                        show_header=_safe_bool(gl, "GroupHeader", True),
                        show_footer=_safe_bool(gl, "GroupFooter", False),
                    )
                )
            except Exception:
                pass
    except Exception:
        pass
    return groupings


def extract_report(
    app: Any,
    report_name: str,
    timeout_seconds: float = 30.0,
) -> tuple[ReportDef | None, list[ExtractionWarning]]:
    """Extract a single ReportDef by opening it in design view."""
    warnings: list[ExtractionWarning] = []

    opened = False
    try:
        opened = _open_report_with_timeout(app, report_name, timeout_seconds)
        if not opened:
            warnings.append(
                ExtractionWarning(
                    object_type="report",
                    object_name=report_name,
                    warning_code="OPEN_TIMEOUT",
                    message=(
                        f"Report '{report_name}' took >{timeout_seconds}s to open. Skipped."
                    ),
                    severity=WarningSeverity.WARNING,
                )
            )
            return None, warnings
    except Exception as e:
        warnings.append(
            ExtractionWarning(
                object_type="report",
                object_name=report_name,
                warning_code="OPEN_FAILED",
                message=f"Could not open report '{report_name}' in design view: {e}",
                severity=WarningSeverity.ERROR,
            )
        )
        return None, warnings

    try:
        report_obj = app.Reports(report_name)
    except Exception as e:
        warnings.append(
            ExtractionWarning(
                object_type="report",
                object_name=report_name,
                warning_code="REPORT_REFERENCE_FAILED",
                message=f"Opened report but couldn't get reference: {e}",
                severity=WarningSeverity.ERROR,
            )
        )
        _safe_close_report(app, report_name)
        return None, warnings

    # Extract sections
    sections: list[ReportSection] = []
    extraction_warnings: list[str] = []

    try:
        # Access reports have numbered sections 0–6
        for section_num in range(7):
            try:
                section = report_obj.Section(section_num)
                if section is None:
                    continue
                section_name = REPORT_SECTION_NAMES.get(section_num, f"Section{section_num}")
                height = _safe_int(section, "Height")

                sec_controls: list[ControlDef] = []
                try:
                    ctrl_count = section.Controls.Count
                    for i in range(ctrl_count):
                        try:
                            ctrl = section.Controls(i)
                            ctrl_def, ctrl_warns = _extract_control(ctrl)
                            if ctrl_def:
                                sec_controls.append(ctrl_def)
                            extraction_warnings.extend(ctrl_warns)
                        except Exception as e:
                            extraction_warnings.append(
                                f"Section {section_name}, control {i}: {e}"
                            )
                except Exception:
                    pass

                sections.append(
                    ReportSection(
                        section_type=section_num,
                        section_name=section_name,
                        controls=sec_controls,
                        height=height,
                    )
                )
            except Exception:
                # Section doesn't exist for this report — that's fine
                pass
    except Exception as e:
        extraction_warnings.append(f"Could not enumerate sections: {e}")

    groupings = _extract_groupings(report_obj)

    has_module = _safe_bool(report_obj, "HasModule")
    vba_module_name = f"Report_{report_name}" if has_module else None

    report = ReportDef(
        name=report_name,
        record_source=_safe_str(report_obj, "RecordSource"),
        sections=sections,
        groupings=groupings,
        order_by=_safe_str(report_obj, "OrderBy"),
        filter=_safe_str(report_obj, "Filter"),
        has_vba_module=has_module,
        vba_module_name=vba_module_name,
        extraction_warnings=extraction_warnings,
    )

    _safe_close_report(app, report_name)
    return report, warnings


def _safe_close_report(app: Any, report_name: str) -> None:
    try:
        app.DoCmd.Close(AC_REPORT, report_name)
    except Exception:
        pass


def extract_all_reports(
    app: Any,
    timeout_seconds: float = 30.0,
    heartbeat_interval: int = 10,
) -> tuple[list[ReportDef], list[ExtractionWarning]]:
    """Extract all reports from CurrentProject.AllReports."""
    reports: list[ReportDef] = []
    all_warnings: list[ExtractionWarning] = []

    try:
        all_reports = app.CurrentProject.AllReports
        report_count = all_reports.Count
    except Exception as e:
        all_warnings.append(
            ExtractionWarning(
                object_type="database",
                object_name="AllReports",
                warning_code="COM_ERROR",
                message=f"Could not access AllReports collection: {e}",
                severity=WarningSeverity.ERROR,
            )
        )
        return reports, all_warnings

    for i in range(report_count):
        report_name = f"<index {i}>"
        try:
            report_name = str(all_reports(i).Name)
        except Exception:
            pass

        # Heartbeat check
        if i > 0 and i % heartbeat_interval == 0:
            try:
                _ = app.Version
            except Exception:
                all_warnings.append(
                    ExtractionWarning(
                        object_type="database",
                        object_name="COMSession",
                        warning_code="SESSION_UNRESPONSIVE",
                        message=f"Access COM session unresponsive after report {i}.",
                        severity=WarningSeverity.ERROR,
                    )
                )
                break

        report_def, warnings = extract_report(app, report_name, timeout_seconds)
        all_warnings.extend(warnings)
        if report_def is not None:
            reports.append(report_def)

        # Reset display state
        try:
            app.Echo(True)
        except Exception:
            pass

    return reports, all_warnings
