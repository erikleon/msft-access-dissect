"""
TableExtractor: extract TableDef objects from a DAO Database via COM.

Skips MSys* system tables by default (--include-system to override).
Handles linked tables by recording connection metadata (no data traversal).
Flags OLE Object and Attachment fields with extraction warnings.
"""

from __future__ import annotations

from typing import Any

from access_dissect.catalog.constants import (
    dao_type_to_field_type,
)
from access_dissect.catalog.models import (
    FieldDef,
    FieldType,
    IndexDef,
    IndexField,
    LinkedTableSource,
    TableDef,
    WarningSeverity,
    ExtractionWarning,
)


# DAO TableDef attribute flags
DB_SYSTEMOBJECT = 0x80000002
DB_HIDDEN = 0x00000001
DB_ATTACHED_TABLE = 0x00000040   # linked table (non-ODBC)
DB_ATTACHED_ODBC = 0x20000000    # linked ODBC table

_BINARY_FIELD_TYPES = {FieldType.OLE_OBJECT, FieldType.ATTACHMENT}


def _strip_credentials(connect_string: str) -> str:
    """
    Remove UID= and PWD= from a connect string to avoid leaking credentials.
    """
    import re
    cleaned = re.sub(r"(?i)(UID|PWD)=[^;]*;?", "", connect_string)
    return cleaned.strip(";").strip()


def _parse_linked_source(table_def: Any) -> LinkedTableSource | None:
    """Parse a linked table's Connect string into a LinkedTableSource model."""
    try:
        connect = str(table_def.Connect or "")
    except Exception:
        return None

    if not connect:
        return None

    cleaned = _strip_credentials(connect)
    lower = connect.lower()

    return LinkedTableSource(
        connect_string=cleaned,
        source_table_name=_safe_str(table_def, "SourceTableName"),
        is_odbc="odbc;" in lower,
        is_excel=".xls" in lower or "excel" in lower,
        is_sharepoint="sharepoint" in lower or ".list" in lower,
        is_text_file="text;" in lower,
        raw_connect_hint=connect[:25] if connect else None,
    )


def _safe_str(obj: Any, attr: str, default: str | None = None) -> str | None:
    """Safely read a COM property that may raise."""
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


def extract_field(field: Any, ordinal: int) -> FieldDef:
    """Build a FieldDef from a DAO Field object."""
    dao_code = _safe_int(field, "Type", 0) or 0
    field_type = dao_type_to_field_type(dao_code)

    # Hyperlinks are stored as Memo fields with a special flag
    try:
        if _safe_bool(field, "AllowZeroLength") and dao_code == 12:
            # Check for hyperlink flag in field properties
            try:
                for prop in field.Properties:
                    if str(prop.Name).lower() == "ishyperlink" and bool(prop.Value):
                        field_type = FieldType.HYPERLINK
                        break
            except Exception:
                pass
    except Exception:
        pass

    return FieldDef(
        name=str(field.Name),
        ordinal=ordinal,
        field_type=field_type,
        dao_type_code=dao_code,
        size=_safe_int(field, "Size"),
        required=_safe_bool(field, "Required"),
        allow_zero_length=_safe_bool(field, "AllowZeroLength"),
        default_value=_safe_str(field, "DefaultValue"),
        validation_rule=_safe_str(field, "ValidationRule"),
        validation_text=_safe_str(field, "ValidationText"),
        description=_safe_str(field, "Description"),
        auto_number=_safe_bool(field, "AutoNumber") or (dao_code == 4 and _safe_bool(field, "Attributes")),
    )


def extract_index(idx: Any) -> IndexDef:
    """Build an IndexDef from a DAO Index object."""
    idx_fields: list[IndexField] = []
    try:
        for f in idx.Fields:
            idx_fields.append(
                IndexField(
                    field_name=str(f.Name),
                    descending=_safe_bool(f, "Attributes"),
                )
            )
    except Exception:
        pass

    return IndexDef(
        name=str(idx.Name),
        fields=idx_fields,
        unique=_safe_bool(idx, "Unique"),
        primary=_safe_bool(idx, "Primary"),
        required=_safe_bool(idx, "Required"),
        ignore_nulls=_safe_bool(idx, "IgnoreNulls"),
        foreign=_safe_bool(idx, "Foreign"),
    )


def extract_table(
    table_def: Any,
    include_system: bool = False,
) -> tuple[TableDef | None, list[ExtractionWarning]]:
    """
    Extract a single TableDef from a DAO TableDef COM object.

    Returns (table_model, warnings).
    Returns (None, warnings) for tables that should be skipped.
    """
    warnings: list[ExtractionWarning] = []
    name = ""
    try:
        name = str(table_def.Name)
    except Exception:
        return None, warnings

    # Skip MSys* system tables (and ~TMP* temp tables) unless requested
    if name.startswith("MSys") or name.startswith("~TMP"):
        if not include_system:
            return None, warnings

    # Determine flags
    attrs = 0
    try:
        attrs = int(table_def.Attributes)
    except Exception:
        pass

    is_system = bool(attrs & DB_SYSTEMOBJECT) or name.startswith("MSys")
    if is_system and not include_system:
        return None, warnings

    is_hidden = bool(attrs & DB_HIDDEN)
    is_linked = bool(attrs & (DB_ATTACHED_TABLE | DB_ATTACHED_ODBC))

    linked_source = _parse_linked_source(table_def) if is_linked else None

    # Extract fields
    fields: list[FieldDef] = []
    try:
        for i, field in enumerate(table_def.Fields):
            try:
                fd = extract_field(field, ordinal=i)
                fields.append(fd)
                # Flag binary field types
                if fd.field_type in _BINARY_FIELD_TYPES:
                    warnings.append(
                        ExtractionWarning(
                            object_type="table",
                            object_name=name,
                            warning_code="BINARY_FIELD",
                            message=(
                                f"Field '{fd.name}' is {fd.field_type.value} — "
                                "binary data cannot be migrated automatically. "
                                "Define a storage strategy (filesystem, S3, or blob)."
                            ),
                            severity=WarningSeverity.WARNING,
                            details={"field_name": fd.name, "field_type": fd.field_type.value},
                        )
                    )
            except Exception as e:
                warnings.append(
                    ExtractionWarning(
                        object_type="table",
                        object_name=name,
                        warning_code="FIELD_EXTRACTION_ERROR",
                        message=f"Could not extract field at ordinal {i}: {e}",
                        severity=WarningSeverity.WARNING,
                    )
                )
    except Exception as e:
        warnings.append(
            ExtractionWarning(
                object_type="table",
                object_name=name,
                warning_code="FIELDS_INACCESSIBLE",
                message=f"Could not enumerate fields: {e}",
                severity=WarningSeverity.ERROR,
            )
        )

    # Extract indexes
    indexes: list[IndexDef] = []
    has_pk = False
    try:
        for idx in table_def.Indexes:
            try:
                idx_def = extract_index(idx)
                indexes.append(idx_def)
                if idx_def.primary:
                    has_pk = True
                    # Mark primary key fields
                    pk_field_names = {f.field_name for f in idx_def.fields}
                    for fd in fields:
                        if fd.name in pk_field_names:
                            fd.is_primary_key = True
            except Exception:
                pass
    except Exception:
        pass

    if not has_pk and fields and not is_linked:
        warnings.append(
            ExtractionWarning(
                object_type="table",
                object_name=name,
                warning_code="NO_PRIMARY_KEY",
                message=(
                    f"Table '{name}' has no primary key. "
                    "This increases migration complexity and may cause data integrity issues."
                ),
                severity=WarningSeverity.WARNING,
            )
        )

    # Check for fields with spaces (naming issues)
    space_fields = [fd.name for fd in fields if " " in fd.name]
    if space_fields:
        warnings.append(
            ExtractionWarning(
                object_type="table",
                object_name=name,
                warning_code="FIELDS_WITH_SPACES",
                message=(
                    f"Table '{name}' has {len(space_fields)} field(s) with spaces in name: "
                    f"{space_fields[:5]}{'...' if len(space_fields) > 5 else ''}. "
                    "Map to snake_case in the target schema."
                ),
                severity=WarningSeverity.INFO,
                details={"fields": space_fields},
            )
        )

    table = TableDef(
        name=name,
        is_system_table=is_system,
        is_hidden=is_hidden,
        is_linked=is_linked,
        linked_source=linked_source,
        fields=fields,
        indexes=indexes,
        description=_safe_str(table_def, "Description"),
    )
    return table, warnings


def extract_all_tables(
    db: Any,
    include_system: bool = False,
) -> tuple[list[TableDef], list[ExtractionWarning]]:
    """
    Extract all tables from CurrentDb().TableDefs.

    Returns (tables, warnings). Per-object errors are recorded as warnings;
    the pipeline never aborts due to a single bad table.
    """
    tables: list[TableDef] = []
    all_warnings: list[ExtractionWarning] = []

    try:
        table_defs = db.TableDefs
    except Exception as e:
        all_warnings.append(
            ExtractionWarning(
                object_type="database",
                object_name="TableDefs",
                warning_code="COM_ERROR",
                message=f"Could not access TableDefs collection: {e}",
                severity=WarningSeverity.ERROR,
            )
        )
        return tables, all_warnings

    count = 0
    try:
        count = table_defs.Count
    except Exception:
        pass

    for i in range(count):
        try:
            td = table_defs(i)
            table, warnings = extract_table(td, include_system=include_system)
            all_warnings.extend(warnings)
            if table is not None:
                tables.append(table)
        except Exception as e:
            name = f"<index {i}>"
            try:
                name = str(table_defs(i).Name)
            except Exception:
                pass
            all_warnings.append(
                ExtractionWarning(
                    object_type="table",
                    object_name=name,
                    warning_code="COM_ERROR",
                    message=f"Unexpected error extracting table: {e}",
                    severity=WarningSeverity.ERROR,
                )
            )

    return tables, all_warnings
