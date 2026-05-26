"""
DAO field type → target dialect SQL type mappings.
"""

from __future__ import annotations

from access_dissect.catalog.models import FieldType

# FieldType → {dialect: sql_type}
_TYPE_MAP: dict[FieldType, dict[str, str]] = {
    FieldType.BOOLEAN:    {"postgres": "BOOLEAN",        "sqlite": "INTEGER",       "mssql": "BIT"},
    FieldType.BYTE:       {"postgres": "SMALLINT",       "sqlite": "INTEGER",       "mssql": "TINYINT"},
    FieldType.INTEGER:    {"postgres": "SMALLINT",       "sqlite": "INTEGER",       "mssql": "SMALLINT"},
    FieldType.LONG:       {"postgres": "INTEGER",        "sqlite": "INTEGER",       "mssql": "INT"},
    FieldType.CURRENCY:   {"postgres": "NUMERIC(19,4)",  "sqlite": "REAL",          "mssql": "MONEY"},
    FieldType.SINGLE:     {"postgres": "REAL",           "sqlite": "REAL",          "mssql": "REAL"},
    FieldType.DOUBLE:     {"postgres": "DOUBLE PRECISION","sqlite": "REAL",         "mssql": "FLOAT"},
    FieldType.DATE:       {"postgres": "TIMESTAMP",      "sqlite": "TEXT",          "mssql": "DATETIME"},
    FieldType.TEXT:       {"postgres": "VARCHAR",        "sqlite": "TEXT",          "mssql": "NVARCHAR"},
    FieldType.MEMO:       {"postgres": "TEXT",           "sqlite": "TEXT",          "mssql": "NVARCHAR(MAX)"},
    FieldType.OLE_OBJECT: {"postgres": "BYTEA",          "sqlite": "BLOB",          "mssql": "VARBINARY(MAX)"},
    FieldType.HYPERLINK:  {"postgres": "TEXT",           "sqlite": "TEXT",          "mssql": "NVARCHAR(MAX)"},
    FieldType.ATTACHMENT: {"postgres": "BYTEA",          "sqlite": "BLOB",          "mssql": "VARBINARY(MAX)"},
    FieldType.GUID:       {"postgres": "UUID",           "sqlite": "TEXT",          "mssql": "UNIQUEIDENTIFIER"},
    FieldType.BIG_INT:    {"postgres": "BIGINT",         "sqlite": "INTEGER",       "mssql": "BIGINT"},
    FieldType.DECIMAL:    {"postgres": "NUMERIC",        "sqlite": "REAL",          "mssql": "DECIMAL"},
    FieldType.UNKNOWN:    {"postgres": "TEXT",           "sqlite": "TEXT",          "mssql": "NVARCHAR(MAX)"},
}


def field_type_to_sql(field_type: FieldType, dialect: str, size: int | None = None) -> str:
    """
    Convert a FieldType to a SQL type string for the given dialect.
    Applies size for VARCHAR types when provided.
    """
    types = _TYPE_MAP.get(field_type, _TYPE_MAP[FieldType.UNKNOWN])
    sql_type = types.get(dialect, types.get("postgres", "TEXT"))

    # Apply size for text types
    if field_type == FieldType.TEXT and size and size > 0:
        if dialect == "postgres":
            return f"VARCHAR({size})"
        if dialect == "mssql":
            return f"NVARCHAR({size})"
        # SQLite doesn't enforce size, but we include it for documentation
        return f"TEXT"

    if field_type == FieldType.DECIMAL and dialect == "postgres":
        return "NUMERIC(28,6)"  # reasonable default for Access Currency/Decimal

    return sql_type
