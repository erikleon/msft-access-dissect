"""
Pydantic v2 catalog models — the primary contract between all layers.

Extractors write to these models.
Renderers read from these models.
The catalog must be fully roundtrippable to/from JSON without information loss.
"""

from __future__ import annotations

from datetime import datetime
from enum import Enum
from typing import Any

from pydantic import BaseModel, Field, model_validator


# ---------------------------------------------------------------------------
# Enumerations
# ---------------------------------------------------------------------------


class ExtractionScope(str, Enum):
    STRUCTURE = "structure"   # schema + VBA code + forms + queries (no data)
    MIGRATION = "migration"   # structure + data migration scripts
    FULL = "full"             # structure + data scripts + data profiling


class FieldType(str, Enum):
    BOOLEAN = "boolean"
    BYTE = "byte"
    INTEGER = "integer"
    LONG = "long"
    CURRENCY = "currency"
    SINGLE = "single"
    DOUBLE = "double"
    DATE = "date"
    TEXT = "text"
    MEMO = "memo"
    OLE_OBJECT = "ole_object"
    HYPERLINK = "hyperlink"
    ATTACHMENT = "attachment"
    GUID = "guid"
    BIG_INT = "big_int"
    DECIMAL = "decimal"
    UNKNOWN = "unknown"


class QueryType(str, Enum):
    SELECT = "select"
    CROSSTAB = "crosstab"
    DELETE = "delete"
    UPDATE = "update"
    APPEND = "append"
    MAKE_TABLE = "make_table"
    DDL = "ddl"
    PASS_THROUGH = "pass_through"
    UNION = "union"
    COMPOUND = "compound"
    UNKNOWN = "unknown"


class VBAModuleType(str, Enum):
    STANDARD = "standard"
    CLASS = "class"
    FORM_MODULE = "form_module"
    REPORT_MODULE = "report_module"


class WarningSeverity(str, Enum):
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"


# ---------------------------------------------------------------------------
# Field / Index / Table models
# ---------------------------------------------------------------------------


class IndexField(BaseModel):
    field_name: str
    descending: bool = False


class IndexDef(BaseModel):
    name: str
    fields: list[IndexField]
    unique: bool = False
    primary: bool = False
    required: bool = False
    ignore_nulls: bool = False
    foreign: bool = False  # part of a Relation constraint index


class FieldDef(BaseModel):
    name: str
    ordinal: int
    field_type: FieldType
    dao_type_code: int = Field(
        description="Raw DAO integer type code for round-trip fidelity"
    )
    size: int | None = None        # max length for Text fields (0 = unlimited for Memo)
    required: bool = False
    allow_zero_length: bool = False
    default_value: str | None = None
    validation_rule: str | None = None
    validation_text: str | None = None
    description: str | None = None
    auto_number: bool = False
    is_primary_key: bool = False
    is_indexed: bool = False


class LinkedTableSource(BaseModel):
    """Metadata for externally linked tables (SQL Server, Excel, SharePoint, etc.)."""
    connect_string: str = Field(
        description="ODBC/OLE DB connect string with credentials stripped"
    )
    source_table_name: str | None = None
    is_odbc: bool = False
    is_excel: bool = False
    is_sharepoint: bool = False
    is_text_file: bool = False
    raw_connect_hint: str | None = Field(
        default=None,
        description="First 20 chars of connect string to indicate driver type without exposing credentials",
    )


class TableDef(BaseModel):
    name: str
    is_system_table: bool = False
    is_hidden: bool = False
    is_linked: bool = False
    linked_source: LinkedTableSource | None = None
    fields: list[FieldDef] = []
    indexes: list[IndexDef] = []
    record_count: int | None = None   # populated in migration/full scope
    row_count_exact: bool = False     # False if estimated or sampled
    description: str | None = None


# ---------------------------------------------------------------------------
# Query models
# ---------------------------------------------------------------------------


class QueryParameter(BaseModel):
    name: str
    data_type: str  # Access parameter type string (e.g., "Text", "Long Integer")


class QueryDef(BaseModel):
    name: str
    query_type: QueryType
    sql_text: str = Field(description="Verbatim Jet/ACE SQL — never modified")
    parameters: list[QueryParameter] = []
    references_queries: list[str] = Field(
        default=[],
        description="Names of other QueryDefs referenced (populated by analyzer)",
    )
    references_tables: list[str] = Field(
        default=[],
        description="Names of TableDefs referenced (populated by analyzer)",
    )
    is_pass_through: bool = False
    pass_through_connect: str | None = None
    returns_records: bool = True
    in_circular_dependency: bool = False
    description: str | None = None


# ---------------------------------------------------------------------------
# Form / Report control models
# ---------------------------------------------------------------------------


class ControlGeometry(BaseModel):
    """Position and size in Access twips (1 inch = 1440 twips)."""
    left: int
    top: int
    width: int
    height: int


class ControlDef(BaseModel):
    name: str
    control_type: int = Field(description="Raw DAO/Access control type integer")
    control_type_name: str = Field(description="Human-readable control type name")
    caption: str | None = None
    control_source: str | None = None   # bound field name
    row_source: str | None = None       # SQL or value list for combo/list boxes
    row_source_type: str | None = None  # "Table/Query", "Value List", "Field List"
    default_value: str | None = None
    geometry: ControlGeometry | None = None
    is_active_x: bool = False
    active_x_prog_id: str | None = None
    subform_source: str | None = None   # source form/report for subform controls
    link_master_fields: str | None = None
    link_child_fields: str | None = None
    tag: str | None = None
    visible: bool = True
    enabled: bool = True
    tab_index: int | None = None
    events: dict[str, str] = Field(
        default={},
        description="Event name → handler ('[ Event Procedure]', macro name, or =Expr())",
    )


# ---------------------------------------------------------------------------
# Form models
# ---------------------------------------------------------------------------


class FormDef(BaseModel):
    name: str
    record_source: str | None = None   # table or query name
    allow_additions: bool = True
    allow_deletions: bool = True
    allow_edits: bool = True
    default_view: int | None = None    # 0=Single, 1=Continuous, 2=Datasheet, etc.
    popup: bool = False
    modal: bool = False
    controls: list[ControlDef] = []
    events: dict[str, str] = Field(
        default={},
        description="Form-level events: {'OnLoad': '[Event Procedure]', ...}",
    )
    has_vba_module: bool = False
    vba_module_name: str | None = None   # "Form_FormName"
    subforms: list[str] = Field(
        default=[],
        description="Source object names of all embedded subform controls",
    )
    description: str | None = None
    extraction_warnings: list[str] = []


# ---------------------------------------------------------------------------
# Report models
# ---------------------------------------------------------------------------


class ReportSection(BaseModel):
    section_type: int    # 0=Detail, 1=FormHeader/ReportHeader, 2=FormFooter/ReportFooter
    section_name: str
    controls: list[ControlDef] = []
    height: int | None = None  # twips


class ReportGrouping(BaseModel):
    field_or_expression: str
    group_on: int | None = None
    group_interval: int | None = None
    keep_together: bool = False
    show_header: bool = True
    show_footer: bool = False


class ReportDef(BaseModel):
    name: str
    record_source: str | None = None
    sections: list[ReportSection] = []
    groupings: list[ReportGrouping] = []
    order_by: str | None = None
    filter: str | None = None
    has_vba_module: bool = False
    vba_module_name: str | None = None
    description: str | None = None
    extraction_warnings: list[str] = []


# ---------------------------------------------------------------------------
# Macro models
# ---------------------------------------------------------------------------


class MacroAction(BaseModel):
    sequence: int
    action_name: str
    condition: str | None = None
    arguments: dict[str, str] = {}
    comment: str | None = None


class MacroDef(BaseModel):
    name: str
    actions: list[MacroAction] = []
    is_embedded: bool = False         # True if stored inside a form/report object
    parent_object: str | None = None  # form/report name for embedded macros
    is_legacy: bool = True            # False for XML-based Access 2007+ macros
    raw_xml: str | None = None        # embedded macro XML (Access 2007+)


# ---------------------------------------------------------------------------
# VBA module models
# ---------------------------------------------------------------------------


class VBAReference(BaseModel):
    """External library reference from the VBA project."""
    name: str
    description: str = ""
    guid: str | None = None
    major: int | None = None
    minor: int | None = None
    full_path: str | None = None
    is_broken: bool = False  # reference points to a missing/unregistered library


class VBAModule(BaseModel):
    name: str
    module_type: VBAModuleType
    source_code: str | None = Field(
        default=None,
        description="Full VBA source code; None if VBA project is password-protected",
    )
    line_count: int | None = None
    declaration_lines: int | None = None
    is_vba_locked: bool = False
    parent_object: str | None = None  # form/report name for code-behind modules


# ---------------------------------------------------------------------------
# Relationship models
# ---------------------------------------------------------------------------


class RelationField(BaseModel):
    child_field: str   # field in the child/many-side table
    parent_field: str  # field in the parent/one-side table


class RelationDef(BaseModel):
    name: str
    parent_table: str
    child_table: str
    fields: list[RelationField]
    enforce_integrity: bool = False
    cascade_update: bool = False
    cascade_delete: bool = False
    one_to_one: bool = False


# ---------------------------------------------------------------------------
# Data profiling models (scope=full only)
# ---------------------------------------------------------------------------


class FieldProfile(BaseModel):
    field_name: str
    null_count: int | None = None
    non_null_count: int | None = None
    null_rate: float | None = None       # 0.0–1.0
    distinct_count: int | None = None
    cardinality_rate: float | None = None  # distinct_count / non_null_count
    min_value: str | None = None
    max_value: str | None = None
    min_length: int | None = None        # for text fields
    max_length: int | None = None
    avg_length: float | None = None
    is_constant: bool = False            # all non-null values are the same
    suspected_type: str | None = Field(
        default=None,
        description="Suspected semantic type mismatch, e.g. 'date_as_text', 'numeric_as_text'",
    )


class TableProfile(BaseModel):
    table_name: str
    row_count: int
    profile_sample_size: int | None = Field(
        default=None,
        description="If set, profile was computed on this many rows (not the full table)",
    )
    field_profiles: list[FieldProfile] = []


# ---------------------------------------------------------------------------
# Warnings / errors
# ---------------------------------------------------------------------------


class ExtractionWarning(BaseModel):
    object_type: str   # "table", "query", "form", "report", "module", "macro"
    object_name: str
    warning_code: str  # e.g. "VBA_LOCKED", "COM_ERROR", "LINKED_TABLE", "BITNESS_MISMATCH"
    message: str
    severity: WarningSeverity = WarningSeverity.WARNING
    details: dict[str, Any] = {}


# ---------------------------------------------------------------------------
# Database-level properties
# ---------------------------------------------------------------------------


class DatabaseProperties(BaseModel):
    title: str | None = None
    subject: str | None = None
    author: str | None = None
    company: str | None = None
    description: str | None = None
    created_date: datetime | None = None
    last_modified: datetime | None = None
    access_version: str | None = None   # e.g. "16.0" (Office 365)
    jet_version: str | None = None      # Jet/ACE engine version
    file_path: str
    file_size_bytes: int | None = None
    is_accdb: bool = True               # False for legacy .mdb
    vba_references: list[VBAReference] = []
    vba_project_locked: bool = False    # True if VBA project is password-protected
    workgroup_db_path: str | None = None  # path to .mdw security file, if any


# ---------------------------------------------------------------------------
# Top-level catalog
# ---------------------------------------------------------------------------


class AccessCatalog(BaseModel):
    """
    The root catalog model. This is the immutable intermediate format —
    written by extractors, read by renderers and analyzers.

    All downstream processing (rendering, analysis, SQL generation) reads
    exclusively from this model. The extraction phase (requiring Access.exe)
    is fully decoupled from all other phases.
    """

    schema_version: str = "1.0"
    extracted_at: datetime
    extracted_by: str = Field(description="Tool version string, e.g. 'access-dissect 0.1.0'")
    scope: ExtractionScope
    properties: DatabaseProperties
    tables: list[TableDef] = []
    queries: list[QueryDef] = []
    forms: list[FormDef] = []
    reports: list[ReportDef] = []
    macros: list[MacroDef] = []
    modules: list[VBAModule] = []
    relations: list[RelationDef] = []
    table_profiles: list[TableProfile] = Field(
        default=[],
        description="Populated only when scope=full",
    )
    warnings: list[ExtractionWarning] = []
    dependency_graph: dict[str, list[str]] = Field(
        default={},
        description="Adjacency list of object dependencies; populated by the analyze step. "
        "Node keys are 'type:name', e.g. 'table:Customers', 'form:frmOrders'.",
    )

    # ---------------------------------------------------------------------------
    # Convenience accessors
    # ---------------------------------------------------------------------------

    def table_names(self) -> set[str]:
        return {t.name for t in self.tables}

    def query_names(self) -> set[str]:
        return {q.name for q in self.queries}

    def form_names(self) -> set[str]:
        return {f.name for f in self.forms}

    def report_names(self) -> set[str]:
        return {r.name for r in self.reports}

    def module_names(self) -> set[str]:
        return {m.name for m in self.modules}

    def warnings_by_severity(self, severity: WarningSeverity) -> list[ExtractionWarning]:
        return [w for w in self.warnings if w.severity == severity]

    def error_count(self) -> int:
        return sum(1 for w in self.warnings if w.severity == WarningSeverity.ERROR)

    def has_vba(self) -> bool:
        return len(self.modules) > 0

    def total_vba_lines(self) -> int:
        return sum(m.line_count or 0 for m in self.modules)

    @model_validator(mode="after")
    def _validate_scope_data(self) -> "AccessCatalog":
        """Warn if table_profiles are present but scope != full."""
        if self.table_profiles and self.scope != ExtractionScope.FULL:
            # Allow it — might be loaded from a manually-created catalog
            pass
        return self
