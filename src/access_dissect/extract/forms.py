"""
FormExtractor: extract FormDef objects by opening each form in design view.

This is the most fragile extraction stage because it requires Access UI operations.
Each form is opened with DoCmd.OpenForm in design view, enumerated, and closed.

Robustness features:
  - Per-form 30-second timeout via threading.Timer
  - Per-object exception isolation
  - Session heartbeat check every 10 forms
  - reset_display_state() after each form
"""

from __future__ import annotations

import threading
from typing import Any

from access_dissect.catalog.constants import control_type_name
from access_dissect.catalog.models import (
    ControlDef,
    ControlGeometry,
    ExtractionWarning,
    FormDef,
    WarningSeverity,
)

# Access constants
AC_FORM = 2
AC_DESIGN = 1


def _safe_str(obj: Any, attr: str, default: str | None = None) -> str | None:
    try:
        val = getattr(obj, attr)
        return str(val) if val is not None else default
    except Exception:
        return default


def _safe_bool(obj: Any, attr: str, default: bool = False) -> bool:
    try:
        return bool(getattr(obj, attr))
    except Exception:
        return default


def _safe_int(obj: Any, attr: str, default: int | None = None) -> int | None:
    try:
        val = getattr(obj, attr)
        return int(val) if val is not None else default
    except Exception:
        return default


_EVENT_PROPERTIES = {
    "OnClick", "OnDblClick", "OnMouseDown", "OnMouseUp", "OnMouseMove",
    "OnKeyDown", "OnKeyUp", "OnKeyPress", "OnChange", "OnUpdated",
    "OnEnter", "OnExit", "OnGotFocus", "OnLostFocus", "OnLoad",
    "OnUnload", "OnOpen", "OnClose", "OnResize", "OnActivate",
    "OnDeactivate", "OnBeforeInsert", "OnAfterInsert", "OnBeforeUpdate",
    "OnAfterUpdate", "OnDelete", "OnBeforeDelConfirm", "OnAfterDelConfirm",
    "OnDirty", "OnNotInList", "OnTimer", "OnCurrent", "OnError",
    "OnFilter", "OnApplyFilter",
}


def _extract_events(obj: Any) -> dict[str, str]:
    """Extract event property bindings from a form or control."""
    events: dict[str, str] = {}
    for event_name in _EVENT_PROPERTIES:
        try:
            val = getattr(obj, event_name, None)
            if val and str(val).strip():
                events[event_name] = str(val).strip()
        except Exception:
            pass
    return events


def _extract_geometry(ctrl: Any) -> ControlGeometry | None:
    """Extract control position/size in twips."""
    try:
        return ControlGeometry(
            left=int(ctrl.Left),
            top=int(ctrl.Top),
            width=int(ctrl.Width),
            height=int(ctrl.Height),
        )
    except Exception:
        return None


def _extract_control(ctrl: Any) -> tuple[ControlDef | None, list[str]]:
    """
    Extract a single ControlDef from an Access control COM object.
    Returns (control, local_warnings).
    """
    local_warnings: list[str] = []
    name = _safe_str(ctrl, "Name", "<unnamed>") or "<unnamed>"
    ctrl_type = _safe_int(ctrl, "ControlType", 0) or 0
    type_name = control_type_name(ctrl_type)

    is_active_x = ctrl_type == 119  # acActiveX
    active_x_prog_id = None
    if is_active_x:
        try:
            active_x_prog_id = str(ctrl.OleObjectSourceDocument or "")
        except Exception:
            pass
        try:
            active_x_prog_id = active_x_prog_id or str(ctrl.OLEClass or "")
        except Exception:
            pass
        local_warnings.append(
            f"ActiveX control '{name}' ({active_x_prog_id or 'unknown ProgID'}) — "
            "may not be reproducible in the target environment."
        )

    # Subform source
    subform_source = None
    link_master = None
    link_child = None
    if ctrl_type == 113:  # acSubform
        subform_source = _safe_str(ctrl, "SourceObject")
        link_master = _safe_str(ctrl, "LinkMasterFields")
        link_child = _safe_str(ctrl, "LinkChildFields")

    events = _extract_events(ctrl)

    control = ControlDef(
        name=name,
        control_type=ctrl_type,
        control_type_name=type_name,
        caption=_safe_str(ctrl, "Caption"),
        control_source=_safe_str(ctrl, "ControlSource"),
        row_source=_safe_str(ctrl, "RowSource"),
        row_source_type=_safe_str(ctrl, "RowSourceType"),
        default_value=_safe_str(ctrl, "DefaultValue"),
        geometry=_extract_geometry(ctrl),
        is_active_x=is_active_x,
        active_x_prog_id=active_x_prog_id if active_x_prog_id else None,
        subform_source=subform_source,
        link_master_fields=link_master,
        link_child_fields=link_child,
        tag=_safe_str(ctrl, "Tag"),
        visible=_safe_bool(ctrl, "Visible", True),
        enabled=_safe_bool(ctrl, "Enabled", True),
        tab_index=_safe_int(ctrl, "TabIndex"),
        events=events,
    )
    return control, local_warnings


def _open_form_with_timeout(
    app: Any,
    form_name: str,
    timeout_seconds: float = 30.0,
) -> bool:
    """
    Open a form in design view with a timeout.
    Returns True if successful, False if timed out or failed.
    """
    success = [False]
    exception_holder: list[Exception] = []

    def _open() -> None:
        try:
            app.DoCmd.OpenForm(form_name, AC_DESIGN)
            success[0] = True
        except Exception as e:
            exception_holder.append(e)

    t = threading.Thread(target=_open, daemon=True)
    t.start()
    t.join(timeout=timeout_seconds)

    if t.is_alive():
        # Timed out — thread is stuck (likely waiting on a dialog)
        return False

    if exception_holder:
        raise exception_holder[0]

    return success[0]


def extract_form(
    app: Any,
    form_name: str,
    timeout_seconds: float = 30.0,
) -> tuple[FormDef | None, list[ExtractionWarning]]:
    """
    Extract a single FormDef by opening it in design view.

    The form is always closed after extraction, even on error.
    """
    warnings: list[ExtractionWarning] = []

    # Attempt to open in design view
    opened = False
    try:
        opened = _open_form_with_timeout(app, form_name, timeout_seconds)
        if not opened:
            warnings.append(
                ExtractionWarning(
                    object_type="form",
                    object_name=form_name,
                    warning_code="OPEN_TIMEOUT",
                    message=(
                        f"Form '{form_name}' took >{timeout_seconds}s to open in design view. "
                        "It may require an ActiveX control or modal dialog. Skipped."
                    ),
                    severity=WarningSeverity.WARNING,
                )
            )
            return None, warnings
    except Exception as e:
        warnings.append(
            ExtractionWarning(
                object_type="form",
                object_name=form_name,
                warning_code="OPEN_FAILED",
                message=f"Could not open form '{form_name}' in design view: {e}",
                severity=WarningSeverity.ERROR,
            )
        )
        return None, warnings

    try:
        form_obj = app.Forms(form_name)
    except Exception as e:
        warnings.append(
            ExtractionWarning(
                object_type="form",
                object_name=form_name,
                warning_code="FORM_REFERENCE_FAILED",
                message=f"Opened form but couldn't get reference: {e}",
                severity=WarningSeverity.ERROR,
            )
        )
        _safe_close_form(app, form_name)
        return None, warnings

    # Extract controls
    controls: list[ControlDef] = []
    subforms: list[str] = []
    extraction_warnings: list[str] = []

    try:
        ctrl_count = form_obj.Controls.Count
        for i in range(ctrl_count):
            try:
                ctrl = form_obj.Controls(i)
                ctrl_def, ctrl_warns = _extract_control(ctrl)
                if ctrl_def:
                    controls.append(ctrl_def)
                    if ctrl_def.subform_source:
                        subforms.append(ctrl_def.subform_source)
                extraction_warnings.extend(ctrl_warns)
            except Exception as e:
                extraction_warnings.append(f"Control at index {i}: {e}")
    except Exception as e:
        extraction_warnings.append(f"Could not enumerate controls: {e}")

    # Form-level properties
    record_source = _safe_str(form_obj, "RecordSource")
    events = _extract_events(form_obj)

    has_module = _safe_bool(form_obj, "HasModule")
    vba_module_name = f"Form_{form_name}" if has_module else None

    form = FormDef(
        name=form_name,
        record_source=record_source,
        allow_additions=_safe_bool(form_obj, "AllowAdditions", True),
        allow_deletions=_safe_bool(form_obj, "AllowDeletions", True),
        allow_edits=_safe_bool(form_obj, "AllowEdits", True),
        default_view=_safe_int(form_obj, "DefaultView"),
        popup=_safe_bool(form_obj, "PopUp"),
        modal=_safe_bool(form_obj, "Modal"),
        controls=controls,
        events=events,
        has_vba_module=has_module,
        vba_module_name=vba_module_name,
        subforms=list(set(subforms)),
        extraction_warnings=extraction_warnings,
    )

    _safe_close_form(app, form_name)
    return form, warnings


def _safe_close_form(app: Any, form_name: str) -> None:
    try:
        app.DoCmd.Close(AC_FORM, form_name)
    except Exception:
        pass


def extract_all_forms(
    app: Any,
    timeout_seconds: float = 30.0,
    heartbeat_interval: int = 10,
) -> tuple[list[FormDef], list[ExtractionWarning]]:
    """
    Extract all forms from CurrentProject.AllForms.

    Opens each form in design view, extracts it, closes it.
    COM session heartbeat check every `heartbeat_interval` forms.
    """
    forms: list[FormDef] = []
    all_warnings: list[ExtractionWarning] = []

    try:
        all_forms = app.CurrentProject.AllForms
        form_count = all_forms.Count
    except Exception as e:
        all_warnings.append(
            ExtractionWarning(
                object_type="database",
                object_name="AllForms",
                warning_code="COM_ERROR",
                message=f"Could not access AllForms collection: {e}",
                severity=WarningSeverity.ERROR,
            )
        )
        return forms, all_warnings

    for i in range(form_count):
        form_name = f"<index {i}>"
        try:
            form_name = str(all_forms(i).Name)
        except Exception:
            pass

        # Heartbeat check every N forms
        if i > 0 and i % heartbeat_interval == 0:
            try:
                _ = app.Version
            except Exception:
                all_warnings.append(
                    ExtractionWarning(
                        object_type="database",
                        object_name="COMSession",
                        warning_code="SESSION_UNRESPONSIVE",
                        message=(
                            f"Access COM session became unresponsive after form {i}. "
                            "Remaining forms will be skipped."
                        ),
                        severity=WarningSeverity.ERROR,
                    )
                )
                break

        form_def, warnings = extract_form(app, form_name, timeout_seconds)
        all_warnings.extend(warnings)
        if form_def is not None:
            forms.append(form_def)

        # Reset display state after each form
        try:
            app.Echo(True)
        except Exception:
            pass

    return forms, all_warnings
