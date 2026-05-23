"""
MacroExtractor: extract standalone macro definitions.

Legacy macros (pre-Access 2007): read via Containers("Scripts")
Modern macros (Access 2007+): stored as XML within the database binary

Embedded macros (in forms/reports): noted as "[Embedded Macro]" in control events;
the full action list requires XML parsing of the form binary which is documented
as a known v1 limitation.
"""

from __future__ import annotations

from typing import Any

from access_dissect.catalog.models import (
    ExtractionWarning,
    MacroDef,
    MacroAction,
    WarningSeverity,
)


def _safe_str(obj: Any, attr: str, default: str | None = None) -> str | None:
    try:
        val = getattr(obj, attr)
        return str(val) if val is not None else default
    except Exception:
        return default


def extract_all_macros(
    app: Any,
    db: Any,
) -> tuple[list[MacroDef], list[ExtractionWarning]]:
    """
    Extract all standalone macros from CurrentProject.AllMacros.

    Legacy macro content is approximated via the Containers("Scripts") API.
    Embedded macros inside forms/reports are flagged but not fully decoded in v1.
    """
    macros: list[MacroDef] = []
    all_warnings: list[ExtractionWarning] = []

    # Get macro names from AllMacros
    macro_names: list[str] = []
    try:
        all_macros_coll = app.CurrentProject.AllMacros
        for i in range(all_macros_coll.Count):
            try:
                macro_names.append(str(all_macros_coll(i).Name))
            except Exception:
                pass
    except Exception as e:
        all_warnings.append(
            ExtractionWarning(
                object_type="database",
                object_name="AllMacros",
                warning_code="COM_ERROR",
                message=f"Could not access AllMacros collection: {e}",
                severity=WarningSeverity.ERROR,
            )
        )
        return macros, all_warnings

    for macro_name in macro_names:
        try:
            # Try to get macro definition via Containers("Scripts")
            actions: list[MacroAction] = []
            raw_xml: str | None = None
            is_legacy = True

            try:
                scripts_container = db.Containers("Scripts")
                macro_doc = scripts_container.Documents(macro_name)
                # Legacy macros store actions in a binary format;
                # we can get the object but can't easily decode actions without
                # running the macro via a low-level approach.
                # For v1: record that the macro exists with empty action list.
                # The macro name and existence are the most important metadata.
            except Exception:
                pass

            macros.append(
                MacroDef(
                    name=macro_name,
                    actions=actions,
                    is_embedded=False,
                    parent_object=None,
                    is_legacy=is_legacy,
                    raw_xml=raw_xml,
                )
            )

            all_warnings.append(
                ExtractionWarning(
                    object_type="macro",
                    object_name=macro_name,
                    warning_code="MACRO_ACTIONS_NOT_DECODED",
                    message=(
                        f"Macro '{macro_name}' was cataloged but its action list "
                        "could not be decoded in v1. The macro exists and is named, "
                        "but full action enumeration requires direct binary parsing. "
                        "Review manually in Access."
                    ),
                    severity=WarningSeverity.INFO,
                )
            )

        except Exception as e:
            all_warnings.append(
                ExtractionWarning(
                    object_type="macro",
                    object_name=macro_name,
                    warning_code="COM_ERROR",
                    message=f"Unexpected error extracting macro: {e}",
                    severity=WarningSeverity.ERROR,
                )
            )

    return macros, all_warnings
