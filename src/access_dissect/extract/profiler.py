"""
DataProfiler: compute null rates, cardinality, min/max, and type hints per field.

scope=full only. Uses pyodbc for all data reads.
Tables > 100k rows are sampled (top 10k rows).
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from access_dissect.catalog.models import (
    ExtractionWarning,
    FieldDef,
    FieldProfile,
    FieldType,
    TableDef,
    TableProfile,
    WarningSeverity,
)

_SKIP_FIELD_TYPES = {FieldType.OLE_OBJECT, FieldType.ATTACHMENT}
_PROFILE_SAMPLE_THRESHOLD = 100_000
_PROFILE_SAMPLE_SIZE = 10_000


def _quote_identifier(name: str) -> str:
    return f"[{name}]"


def _build_connection_string(file_path: Path, password: str | None) -> str:
    path_str = str(file_path)
    conn = f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={path_str};"
    if password:
        conn += f"PWD={password};"
    return conn


def _profile_field(
    cursor: Any,
    table_name: str,
    field: FieldDef,
    row_count: int,
    is_text: bool,
) -> FieldProfile:
    """Compute profiling stats for a single field using aggregate SQL."""
    tbl = _quote_identifier(table_name)
    col = _quote_identifier(field.name)
    profile = FieldProfile(field_name=field.name)

    try:
        # Null count
        cursor.execute(f"SELECT COUNT(*) FROM {tbl} WHERE {col} IS NULL")
        row = cursor.fetchone()
        null_count = int(row[0]) if row else 0
        profile.null_count = null_count
        profile.non_null_count = row_count - null_count
        profile.null_rate = null_count / row_count if row_count > 0 else 0.0
        profile.is_constant = profile.non_null_count <= 1
    except Exception:
        pass

    try:
        # Distinct count
        cursor.execute(f"SELECT COUNT(DISTINCT {col}) FROM {tbl}")
        row = cursor.fetchone()
        if row and row[0] is not None:
            distinct = int(row[0])
            profile.distinct_count = distinct
            non_null = profile.non_null_count or 0
            profile.cardinality_rate = distinct / non_null if non_null > 0 else 0.0
            if distinct == 1 and non_null > 0:
                profile.is_constant = True
    except Exception:
        pass

    # Min/max (for non-binary fields)
    if field.field_type not in _SKIP_FIELD_TYPES:
        try:
            cursor.execute(f"SELECT MIN({col}), MAX({col}) FROM {tbl}")
            row = cursor.fetchone()
            if row:
                profile.min_value = str(row[0]) if row[0] is not None else None
                profile.max_value = str(row[1]) if row[1] is not None else None
        except Exception:
            pass

    # Length stats for text fields
    if is_text:
        try:
            cursor.execute(
                f"SELECT MIN(LEN({col})), MAX(LEN({col})), AVG(LEN({col})) FROM {tbl} "
                f"WHERE {col} IS NOT NULL"
            )
            row = cursor.fetchone()
            if row and row[0] is not None:
                profile.min_length = int(row[0])
                profile.max_length = int(row[1])
                profile.avg_length = float(row[2]) if row[2] is not None else None
        except Exception:
            pass

        # Heuristic: detect dates stored as text (pattern: ##/##/####)
        if profile.min_value and profile.max_value:
            date_pattern = re.compile(r"^\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}$")
            if date_pattern.match(str(profile.min_value)):
                profile.suspected_type = "date_as_text"

    # Heuristic: detect numbers stored as text
    if is_text and profile.min_value and profile.max_value:
        try:
            float(profile.min_value)
            float(profile.max_value)
            if profile.suspected_type is None:
                profile.suspected_type = "numeric_as_text"
        except (ValueError, TypeError):
            pass

    return profile


def profile_all_tables(
    file_path: Path,
    tables: list[TableDef],
    password: str | None = None,
) -> tuple[list[TableProfile], list[ExtractionWarning]]:
    """
    Profile all non-linked, non-system tables.

    Tables over PROFILE_SAMPLE_THRESHOLD rows are sampled.
    Returns (table_profiles, warnings).
    """
    profiles: list[TableProfile] = []
    warnings: list[ExtractionWarning] = []

    try:
        import pyodbc
    except ImportError:
        warnings.append(
            ExtractionWarning(
                object_type="database",
                object_name="DataProfiler",
                warning_code="PYODBC_NOT_INSTALLED",
                message="pyodbc not installed; data profiling not available.",
                severity=WarningSeverity.WARNING,
            )
        )
        return profiles, warnings

    conn_str = _build_connection_string(file_path, password)
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
    except Exception as e:
        warnings.append(
            ExtractionWarning(
                object_type="database",
                object_name="DataProfiler",
                warning_code="ODBC_CONNECT_FAILED",
                message=f"Could not connect for profiling: {e}",
                severity=WarningSeverity.ERROR,
            )
        )
        return profiles, warnings

    cursor = conn.cursor()

    for table in tables:
        if table.is_linked or table.is_system_table:
            continue

        row_count = table.record_count or 0
        sample_size: int | None = None

        # Determine actual row count if we don't have it
        if row_count == 0:
            try:
                cursor.execute(
                    f"SELECT COUNT(*) FROM {_quote_identifier(table.name)}"
                )
                result = cursor.fetchone()
                row_count = int(result[0]) if result else 0
            except Exception:
                pass

        if row_count == 0:
            continue

        # For very large tables, note we're sampling
        if row_count > _PROFILE_SAMPLE_THRESHOLD:
            sample_size = _PROFILE_SAMPLE_SIZE
            warnings.append(
                ExtractionWarning(
                    object_type="table",
                    object_name=table.name,
                    warning_code="PROFILE_SAMPLED",
                    message=(
                        f"Table '{table.name}' has {row_count:,} rows. "
                        f"Profiling computed on a sample of {_PROFILE_SAMPLE_SIZE:,} rows."
                    ),
                    severity=WarningSeverity.INFO,
                )
            )

        field_profiles: list[FieldProfile] = []
        for field in table.fields:
            if field.field_type in _SKIP_FIELD_TYPES:
                continue
            is_text = field.field_type in {FieldType.TEXT, FieldType.MEMO, FieldType.HYPERLINK}
            try:
                fp = _profile_field(cursor, table.name, field, row_count, is_text)
                field_profiles.append(fp)

                # High-null field warning
                if fp.null_rate is not None and fp.null_rate > 0.8 and row_count > 10:
                    warnings.append(
                        ExtractionWarning(
                            object_type="table",
                            object_name=table.name,
                            warning_code="HIGH_NULL_RATE",
                            message=(
                                f"Field '{table.name}.{field.name}' is "
                                f"{fp.null_rate*100:.0f}% NULL. "
                                "Consider whether this field is still needed."
                            ),
                            severity=WarningSeverity.INFO,
                            details={
                                "field": field.name,
                                "null_rate": round(fp.null_rate, 3),
                            },
                        )
                    )
            except Exception as e:
                warnings.append(
                    ExtractionWarning(
                        object_type="table",
                        object_name=table.name,
                        warning_code="FIELD_PROFILE_ERROR",
                        message=f"Could not profile field '{field.name}': {e}",
                        severity=WarningSeverity.WARNING,
                    )
                )

        profiles.append(
            TableProfile(
                table_name=table.name,
                row_count=row_count,
                profile_sample_size=sample_size,
                field_profiles=field_profiles,
            )
        )

    cursor.close()
    conn.close()
    return profiles, warnings
