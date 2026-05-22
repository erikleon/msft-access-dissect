"""
QueryExtractor: extract QueryDef objects from a DAO Database via COM.

SQL text is always stored verbatim (Jet/ACE dialect).
Translation to target dialects happens in the render phase.
"""

from __future__ import annotations

from typing import Any

from access_dissect.catalog.constants import dao_query_type_to_enum
from access_dissect.catalog.models import (
    ExtractionWarning,
    QueryDef,
    QueryParameter,
    QueryType,
    WarningSeverity,
)


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


def _safe_int(obj: Any, attr: str, default: int = 0) -> int:
    try:
        return int(getattr(obj, attr))
    except Exception:
        return default


def extract_query(query_def: Any) -> tuple[QueryDef | None, list[ExtractionWarning]]:
    """
    Extract a single QueryDef from a DAO QueryDef COM object.

    Returns (query_model, warnings).
    System/temp queries (names starting with ~) are skipped.
    """
    warnings: list[ExtractionWarning] = []
    name = ""
    try:
        name = str(query_def.Name)
    except Exception:
        return None, warnings

    # Skip temporary queries
    if name.startswith("~"):
        return None, warnings

    sql_text = _safe_str(query_def, "SQL", "") or ""
    dao_type = _safe_int(query_def, "Type", 0)
    query_type = dao_query_type_to_enum(dao_type)

    # Pass-through query metadata
    is_pass_through = query_type == QueryType.PASS_THROUGH
    pass_through_connect = None
    if is_pass_through:
        pass_through_connect = _safe_str(query_def, "Connect")
        warnings.append(
            ExtractionWarning(
                object_type="query",
                object_name=name,
                warning_code="PASS_THROUGH_QUERY",
                message=(
                    f"Query '{name}' is a pass-through query. "
                    "It contains server-side SQL (likely T-SQL) that may not be portable. "
                    "Manual review required during migration."
                ),
                severity=WarningSeverity.WARNING,
                details={"connect_hint": (pass_through_connect or "")[:30]},
            )
        )

    returns_records = _safe_bool(query_def, "ReturnsRecords", default=True)

    # Extract parameters
    parameters: list[QueryParameter] = []
    try:
        params = query_def.Parameters
        for p in params:
            try:
                parameters.append(
                    QueryParameter(
                        name=str(p.Name),
                        data_type=str(p.Type),
                    )
                )
            except Exception:
                pass
    except Exception:
        pass

    query = QueryDef(
        name=name,
        query_type=query_type,
        sql_text=sql_text,
        parameters=parameters,
        is_pass_through=is_pass_through,
        pass_through_connect=pass_through_connect,
        returns_records=returns_records,
    )
    return query, warnings


def extract_all_queries(db: Any) -> tuple[list[QueryDef], list[ExtractionWarning]]:
    """
    Extract all queries from CurrentDb().QueryDefs.

    Per-object errors are isolated; the pipeline continues on failure.
    """
    queries: list[QueryDef] = []
    all_warnings: list[ExtractionWarning] = []

    try:
        query_defs = db.QueryDefs
    except Exception as e:
        all_warnings.append(
            ExtractionWarning(
                object_type="database",
                object_name="QueryDefs",
                warning_code="COM_ERROR",
                message=f"Could not access QueryDefs collection: {e}",
                severity=WarningSeverity.ERROR,
            )
        )
        return queries, all_warnings

    count = 0
    try:
        count = query_defs.Count
    except Exception:
        pass

    for i in range(count):
        name = f"<index {i}>"
        try:
            qd = query_defs(i)
            try:
                name = str(qd.Name)
            except Exception:
                pass
            query, warnings = extract_query(qd)
            all_warnings.extend(warnings)
            if query is not None:
                queries.append(query)
        except Exception as e:
            all_warnings.append(
                ExtractionWarning(
                    object_type="query",
                    object_name=name,
                    warning_code="COM_ERROR",
                    message=f"Unexpected error extracting query: {e}",
                    severity=WarningSeverity.ERROR,
                )
            )

    return queries, all_warnings
