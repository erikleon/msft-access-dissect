"""
DataExtractor: extract row counts and generate INSERT scripts.

Uses pyodbc (ACE ODBC driver) — NOT COM — for all data reads.
This avoids COM memory issues with large result sets.

Scope: migration | full
Skips: linked tables, OLE Object fields, Attachment fields.
"""

from __future__ import annotations

import io
import re
from pathlib import Path
from typing import Any, Generator

from access_dissect.catalog.models import (
    ExtractionWarning,
    FieldType,
    TableDef,
    WarningSeverity,
)

_SKIP_FIELD_TYPES = {FieldType.OLE_OBJECT, FieldType.ATTACHMENT}


def _build_odbc_connection_string(file_path: Path, password: str | None = None) -> str:
    path_str = str(file_path).replace("\\", "\\\\")
    conn = f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={path_str};"
    if password:
        conn += f"PWD={password};"
    return conn


def _quote_identifier(name: str) -> str:
    return f"[{name}]"


def _escape_value(val: Any) -> str:
    """Convert a Python value to an SQL literal."""
    if val is None:
        return "NULL"
    if isinstance(val, bool):
        return "TRUE" if val else "FALSE"
    if isinstance(val, (int, float)):
        return str(val)
    if isinstance(val, bytes):
        return "NULL"  # skip binary data
    # String: escape single quotes
    escaped = str(val).replace("'", "''")
    return f"'{escaped}'"


def get_row_count(cursor: Any, table_name: str) -> int | None:
    """Get an exact row count for a table."""
    try:
        cursor.execute(f"SELECT COUNT(*) FROM {_quote_identifier(table_name)}")
        row = cursor.fetchone()
        return int(row[0]) if row else None
    except Exception:
        return None


def _get_extractable_columns(table: TableDef) -> list[str]:
    """Return field names that can be extracted (skip binary types)."""
    return [
        f.name for f in table.fields
        if f.field_type not in _SKIP_FIELD_TYPES
    ]


def update_row_counts(
    file_path: Path,
    tables: list[TableDef],
    password: str | None = None,
) -> list[ExtractionWarning]:
    """
    Connect via ODBC and update record_count on each non-linked table.
    Modifies tables in-place.
    """
    warnings: list[ExtractionWarning] = []

    try:
        import pyodbc
    except ImportError:
        warnings.append(
            ExtractionWarning(
                object_type="database",
                object_name="DataExtractor",
                warning_code="PYODBC_NOT_INSTALLED",
                message="pyodbc is not installed; row counts not available. Run: pip install pyodbc",
                severity=WarningSeverity.WARNING,
            )
        )
        return warnings

    conn_str = _build_odbc_connection_string(file_path, password)
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
    except Exception as e:
        warnings.append(
            ExtractionWarning(
                object_type="database",
                object_name="DataExtractor",
                warning_code="ODBC_CONNECT_FAILED",
                message=f"Could not connect via ODBC to get row counts: {e}",
                severity=WarningSeverity.WARNING,
            )
        )
        return warnings

    cursor = conn.cursor()
    for table in tables:
        if table.is_linked or table.is_system_table:
            continue
        try:
            count = get_row_count(cursor, table.name)
            if count is not None:
                table.record_count = count
                table.row_count_exact = True
        except Exception as e:
            warnings.append(
                ExtractionWarning(
                    object_type="table",
                    object_name=table.name,
                    warning_code="ROW_COUNT_FAILED",
                    message=f"Could not get row count: {e}",
                    severity=WarningSeverity.INFO,
                )
            )

    cursor.close()
    conn.close()
    return warnings


def generate_insert_scripts(
    file_path: Path,
    tables: list[TableDef],
    output_dir: Path,
    max_rows: int = 0,
    password: str | None = None,
    dialect: str = "generic",
) -> list[ExtractionWarning]:
    """
    Generate INSERT scripts for all non-linked, non-system tables.

    Streams rows in batches of 500 to avoid memory issues.
    Output: one .sql file per table in output_dir/inserts/.

    max_rows=0 means all rows.
    """
    warnings: list[ExtractionWarning] = []

    try:
        import pyodbc
    except ImportError:
        warnings.append(
            ExtractionWarning(
                object_type="database",
                object_name="DataExtractor",
                warning_code="PYODBC_NOT_INSTALLED",
                message="pyodbc is not installed; INSERT scripts not generated.",
                severity=WarningSeverity.WARNING,
            )
        )
        return warnings

    insert_dir = output_dir / "inserts"
    insert_dir.mkdir(parents=True, exist_ok=True)

    conn_str = _build_odbc_connection_string(file_path, password)
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
    except Exception as e:
        warnings.append(
            ExtractionWarning(
                object_type="database",
                object_name="DataExtractor",
                warning_code="ODBC_CONNECT_FAILED",
                message=f"Could not connect via ODBC for INSERT generation: {e}",
                severity=WarningSeverity.ERROR,
            )
        )
        return warnings

    cursor = conn.cursor()

    for table in tables:
        if table.is_linked or table.is_system_table:
            continue

        columns = _get_extractable_columns(table)
        if not columns:
            continue

        try:
            col_clause = ", ".join(_quote_identifier(c) for c in columns)
            sql = f"SELECT {col_clause} FROM {_quote_identifier(table.name)}"
            if max_rows > 0:
                sql = f"SELECT TOP {max_rows} {col_clause} FROM {_quote_identifier(table.name)}"

            cursor.execute(sql)

            out_file = insert_dir / f"{re.sub(r'[^\\w]', '_', table.name)}.sql"
            col_names_sql = ", ".join(_quote_identifier(c) for c in columns)
            table_ref = _quote_identifier(table.name)

            with out_file.open("w", encoding="utf-8") as fh:
                fh.write(f"-- INSERT script for table: {table.name}\n")
                fh.write(f"-- Generated by access-dissect\n\n")

                batch_size = 500
                rows = cursor.fetchmany(batch_size)
                while rows:
                    for row in rows:
                        values = ", ".join(_escape_value(v) for v in row)
                        fh.write(
                            f"INSERT INTO {table_ref} ({col_names_sql}) VALUES ({values});\n"
                        )
                    rows = cursor.fetchmany(batch_size)

        except Exception as e:
            warnings.append(
                ExtractionWarning(
                    object_type="table",
                    object_name=table.name,
                    warning_code="INSERT_SCRIPT_FAILED",
                    message=f"Could not generate INSERT script: {e}",
                    severity=WarningSeverity.WARNING,
                )
            )

    cursor.close()
    conn.close()
    return warnings
