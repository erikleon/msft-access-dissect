"""
DAO constants, control type enums, and lookup tables.

These values are sourced from the DAO 3.6 / ACE type library.
"""

from access_dissect.catalog.models import FieldType

# ---------------------------------------------------------------------------
# DAO field type code → FieldType enum
# Source: DAO 3.6 DataTypeEnum
# ---------------------------------------------------------------------------

DAO_TYPE_MAP: dict[int, FieldType] = {
    1: FieldType.BOOLEAN,       # dbBoolean
    2: FieldType.BYTE,          # dbByte
    3: FieldType.INTEGER,       # dbInteger (16-bit)
    4: FieldType.LONG,          # dbLong (32-bit)
    5: FieldType.CURRENCY,      # dbCurrency
    6: FieldType.SINGLE,        # dbSingle
    7: FieldType.DOUBLE,        # dbDouble
    8: FieldType.DATE,          # dbDate
    10: FieldType.TEXT,         # dbText
    11: FieldType.OLE_OBJECT,   # dbLongBinary (OLE Object)
    12: FieldType.MEMO,         # dbMemo
    15: FieldType.GUID,         # dbGUID
    16: FieldType.BIG_INT,      # dbBigInt (Access 2016+)
    17: FieldType.ATTACHMENT,   # dbAttachment (Access 2007+)
    20: FieldType.DECIMAL,      # dbDecimal
    101: FieldType.HYPERLINK,   # Custom hyperlink (stored as Memo + hyperlink flag)
}


def dao_type_to_field_type(dao_code: int) -> FieldType:
    """Map a raw DAO type integer to a FieldType enum value."""
    return DAO_TYPE_MAP.get(dao_code, FieldType.UNKNOWN)


# ---------------------------------------------------------------------------
# DAO query type code → QueryType enum
# Source: DAO 3.6 QueryTypeEnum
# ---------------------------------------------------------------------------

from access_dissect.catalog.models import QueryType  # noqa: E402

DAO_QUERY_TYPE_MAP: dict[int, QueryType] = {
    0: QueryType.SELECT,
    1: QueryType.CROSSTAB,
    2: QueryType.DELETE,
    3: QueryType.UPDATE,
    4: QueryType.APPEND,
    5: QueryType.MAKE_TABLE,
    6: QueryType.DDL,
    7: QueryType.PASS_THROUGH,
    8: QueryType.UNION,
    9: QueryType.COMPOUND,
}


def dao_query_type_to_enum(dao_code: int) -> QueryType:
    return DAO_QUERY_TYPE_MAP.get(dao_code, QueryType.UNKNOWN)


# ---------------------------------------------------------------------------
# Access control type codes → human-readable names
# Source: Access AcControlType enum
# ---------------------------------------------------------------------------

CONTROL_TYPE_NAMES: dict[int, str] = {
    100: "Label",
    101: "Rectangle",
    102: "Line",
    103: "Image",
    104: "CommandButton",
    106: "OptionButton",
    107: "CheckBox",
    108: "OptionGroup",
    109: "BoundObjectFrame",
    110: "TextBox",
    111: "ListBox",
    112: "ComboBox",
    113: "Subform",
    114: "UnboundObjectFrame",
    118: "PageBreak",
    119: "ActiveX",
    120: "ToggleButton",
    123: "TabControl",
    124: "Page",
    126: "NavigationControl",
    127: "NavigationButton",
}


def control_type_name(type_code: int) -> str:
    return CONTROL_TYPE_NAMES.get(type_code, f"Unknown({type_code})")


# ---------------------------------------------------------------------------
# DAO Relation attribute bit flags
# Source: DAO 3.6 RelationAttributeEnum
# ---------------------------------------------------------------------------

DB_REL_UNIQUE = 1          # One-to-one relationship
DB_REL_DONTENFORCE = 2     # Referential integrity NOT enforced (flag is inverted)
DB_REL_INHERITFIELDS = 4   # (legacy, ignore)
DB_REL_LEFT = 16777216     # Left join (internal)
DB_REL_RIGHT = 33554432    # Right join (internal)
DB_CASCADE_UPDATE = 256    # Cascade updates
DB_CASCADE_DELETE = 4096   # Cascade deletes


def parse_relation_attributes(attributes: int) -> dict[str, bool]:
    """Decode the DAO Relation.Attributes bitmask into a plain dict."""
    return {
        "one_to_one": bool(attributes & DB_REL_UNIQUE),
        "enforce_integrity": not bool(attributes & DB_REL_DONTENFORCE),
        "cascade_update": bool(attributes & DB_CASCADE_UPDATE),
        "cascade_delete": bool(attributes & DB_CASCADE_DELETE),
    }


# ---------------------------------------------------------------------------
# Access form default-view codes
# ---------------------------------------------------------------------------

FORM_DEFAULT_VIEW: dict[int, str] = {
    0: "Single Form",
    1: "Continuous Forms",
    2: "Datasheet",
    3: "PivotTable",
    4: "PivotChart",
    6: "Split Form",
}

# ---------------------------------------------------------------------------
# Report section type codes
# ---------------------------------------------------------------------------

REPORT_SECTION_NAMES: dict[int, str] = {
    0: "Detail",
    1: "Form Header / Report Header",
    2: "Form Footer / Report Footer",
    3: "Page Header",
    4: "Page Footer",
    5: "Group Header",
    6: "Group Footer",
}
