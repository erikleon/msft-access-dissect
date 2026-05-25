"""
ComplexityScorer: compute a migration complexity score (1–10) for each object.

Weighted rule violations summed per object, capped at 10.
Higher score = more difficult to migrate.

Score bands:
  1–3: Low    — straightforward migration
  4–6: Medium — some manual work required
  7–10: High  — significant complexity, plan carefully
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from access_dissect.catalog.models import (
    AccessCatalog,
    FieldType,
    FormDef,
    QueryDef,
    QueryType,
    ReportDef,
    TableDef,
    VBAModule,
)

# Access-specific SQL function names that require manual rewriting
_DOMAIN_AGGREGATE_PATTERN = re.compile(
    r"\b(DLookup|DSum|DCount|DAvg|DMax|DMin|DFirst|DLast|DStDev|DVar)\s*\(",
    re.IGNORECASE,
)
_ACCESS_FUNCTION_PATTERN = re.compile(
    r"\b(IIF|NZ|Format|DatePart|DateDiff|DateAdd|Choose|Switch)\s*\(",
    re.IGNORECASE,
)
_ON_ERROR_RESUME = re.compile(r"On\s+Error\s+Resume\s+Next", re.IGNORECASE)
_DAO_ADO_PATTERN = re.compile(
    r"\b(CurrentDb|OpenRecordset|CreateRecordset|ADODB|DAO\.Recordset|rs\.Open)\b",
    re.IGNORECASE,
)

_BINARY_FIELD_TYPES = {FieldType.OLE_OBJECT, FieldType.ATTACHMENT}

# External COM libraries that indicate portability issues
_COMPLEX_VBA_LIBS = {
    "excel", "word", "outlook", "powerpoint", "msproject",
    "visio", "access", "msoffice",
}


@dataclass
class ObjectScore:
    object_type: str
    object_name: str
    score: float
    reasons: list[str] = field(default_factory=list)

    @property
    def band(self) -> str:
        if self.score <= 3:
            return "low"
        if self.score <= 6:
            return "medium"
        return "high"


def score_table(table: TableDef) -> ObjectScore:
    score = 0.0
    reasons: list[str] = []

    # Binary fields
    binary_fields = [f for f in table.fields if f.field_type in _BINARY_FIELD_TYPES]
    if binary_fields:
        score += 1.0 * len(binary_fields)
        reasons.append(
            f"{len(binary_fields)} binary field(s) (OLE/Attachment): "
            f"{[f.name for f in binary_fields[:3]]}"
        )

    # No primary key
    has_pk = any(f.is_primary_key for f in table.fields) or any(
        idx.primary for idx in table.indexes
    )
    if not has_pk and table.fields:
        score += 1.0
        reasons.append("No primary key defined")

    # Fields with spaces
    space_fields = [f.name for f in table.fields if " " in f.name]
    if space_fields:
        score += 0.5 * min(len(space_fields), 4)
        reasons.append(f"{len(space_fields)} field(s) with spaces in name")

    # Validation rules with Access syntax
    rule_fields = [f for f in table.fields if f.validation_rule]
    if rule_fields:
        score += 0.3 * min(len(rule_fields), 5)
        reasons.append(
            f"{len(rule_fields)} field(s) with Access-specific validation rules"
        )

    # Linked table
    if table.is_linked:
        score += 1.0
        reasons.append(f"Linked table (external source: {table.linked_source and table.linked_source.connect_string[:30]}...)")

    return ObjectScore(
        object_type="table",
        object_name=table.name,
        score=min(score, 10.0),
        reasons=reasons,
    )


def score_query(query: QueryDef) -> ObjectScore:
    score = 0.0
    reasons: list[str] = []
    sql = query.sql_text

    # Pass-through queries (server-side SQL)
    if query.is_pass_through:
        score += 2.0
        reasons.append("Pass-through query (server-side SQL, not portable)")

    # Domain aggregates
    domain_matches = _DOMAIN_AGGREGATE_PATTERN.findall(sql)
    if domain_matches:
        unique = list(set(domain_matches))
        score += 1.5 * min(len(unique), 4)
        reasons.append(
            f"Domain aggregate function(s): {unique[:5]} — "
            "no direct SQL equivalent, needs correlated subquery"
        )

    # Access-specific functions
    func_matches = _ACCESS_FUNCTION_PATTERN.findall(sql)
    if func_matches:
        unique = list(set(func_matches))
        score += 1.0 * min(len(unique), 3)
        reasons.append(f"Access-specific function(s): {unique[:5]}")

    # Circular dependency
    if query.in_circular_dependency:
        score += 1.0
        reasons.append("Part of a circular query dependency chain")

    # Parameters
    if query.parameters:
        score += 0.5 * min(len(query.parameters), 4)
        reasons.append(f"{len(query.parameters)} parameter(s) — need ORM or prepared statement mapping")

    # Action queries used as workflows
    if query.query_type in (QueryType.DELETE, QueryType.UPDATE, QueryType.APPEND, QueryType.MAKE_TABLE):
        score += 0.5
        reasons.append(f"Action query ({query.query_type.value}) — may represent workflow logic")

    return ObjectScore(
        object_type="query",
        object_name=query.name,
        score=min(score, 10.0),
        reasons=reasons,
    )


def score_form(form: FormDef, modules: list[VBAModule]) -> ObjectScore:
    score = 0.0
    reasons: list[str] = []

    # Find associated module
    module_loc = None
    if form.vba_module_name:
        for m in modules:
            if m.name == form.vba_module_name:
                module_loc = m
                break

    if module_loc and module_loc.line_count:
        if module_loc.line_count > 200:
            score += 3.0
            reasons.append(
                f"VBA module has {module_loc.line_count} lines — substantial logic to migrate"
            )
        elif module_loc.line_count > 50:
            score += 1.5
            reasons.append(f"VBA module has {module_loc.line_count} lines")

    # ActiveX controls
    activex_count = sum(1 for c in form.controls if c.is_active_x)
    if activex_count:
        score += 1.0 * activex_count
        prog_ids = list({c.active_x_prog_id for c in form.controls if c.is_active_x and c.active_x_prog_id})
        reasons.append(f"{activex_count} ActiveX control(s): {prog_ids[:3]}")

    # Subforms (nesting complexity)
    if form.subforms:
        score += 0.5 * len(form.subforms)
        reasons.append(f"{len(form.subforms)} subform(s) — nested form complexity")

    # Modal / popup forms need dialog framework
    if form.modal or form.popup:
        score += 0.5
        reasons.append("Modal/popup form — needs dialog framework equivalent")

    return ObjectScore(
        object_type="form",
        object_name=form.name,
        score=min(score, 10.0),
        reasons=reasons,
    )


def score_vba_module(module: VBAModule, vba_references: list[Any]) -> ObjectScore:
    """Score a VBA module for migration complexity."""
    score = 0.0
    reasons: list[str] = []

    if module.is_vba_locked:
        score = 10.0
        reasons.append("VBA project is locked — cannot assess code complexity")
        return ObjectScore(
            object_type="module",
            object_name=module.name,
            score=10.0,
            reasons=reasons,
        )

    code = module.source_code or ""

    # Broken references (already caught at DB level, but reflect per-module)
    broken = [r for r in vba_references if r.is_broken]
    if broken:
        score += 1.0 * min(len(broken), 3)
        reasons.append(f"Database has {len(broken)} broken VBA reference(s)")

    # External COM libraries
    complex_refs = [
        r for r in vba_references
        if any(lib in r.name.lower() for lib in _COMPLEX_VBA_LIBS)
        and not r.is_broken
    ]
    if complex_refs:
        score += 1.5 * min(len(complex_refs), 3)
        reasons.append(f"Uses external COM libraries: {[r.name for r in complex_refs[:4]]}")

    # On Error Resume Next — hidden bugs
    oern_count = len(_ON_ERROR_RESUME.findall(code))
    if oern_count:
        score += 1.0 * min(oern_count, 3)
        reasons.append(
            f"{oern_count} 'On Error Resume Next' statement(s) — hidden error handling"
        )

    # DAO/ADO SQL patterns — potential SQL injection
    dao_matches = _DAO_ADO_PATTERN.findall(code)
    if dao_matches:
        score += 0.5 * min(len(dao_matches), 4)
        reasons.append(f"Direct DAO/ADO database access ({len(dao_matches)} occurrences) — review for SQL injection")

    return ObjectScore(
        object_type="module",
        object_name=module.name,
        score=min(score, 10.0),
        reasons=reasons,
    )


from typing import Any  # noqa: E402 (needed for vba_references type hint)


def score_all_objects(catalog: AccessCatalog) -> list[ObjectScore]:
    """Score all objects in the catalog and return the full list."""
    scores: list[ObjectScore] = []
    vba_refs = catalog.properties.vba_references

    for table in catalog.tables:
        if not table.is_system_table and not table.is_hidden:
            scores.append(score_table(table))

    for query in catalog.queries:
        scores.append(score_query(query))

    for form in catalog.forms:
        scores.append(score_form(form, catalog.modules))

    for module in catalog.modules:
        scores.append(score_vba_module(module, vba_refs))

    return scores
