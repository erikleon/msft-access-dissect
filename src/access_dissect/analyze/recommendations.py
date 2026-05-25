"""
ModernizationAdvisor: produce categorized recommendations from a catalog.

Categories:
  SCHEMA_ISSUES, QUERY_COMPLEXITY, VBA_MIGRATION,
  DATA_QUALITY (scope=full), MIGRATION_BLOCKERS, QUICK_WINS
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from access_dissect.catalog.models import (
    AccessCatalog,
    ExtractionScope,
    FieldType,
    QueryType,
    WarningSeverity,
)

_DOMAIN_AGGREGATE_RE = re.compile(
    r"\b(DLookup|DSum|DCount|DAvg|DMax|DMin|DFirst|DLast)\s*\(", re.IGNORECASE
)
_ON_ERROR_RESUME_RE = re.compile(r"On\s+Error\s+Resume\s+Next", re.IGNORECASE)
_DAO_ADO_RE = re.compile(
    r"\b(CurrentDb|OpenRecordset|CreateRecordset|ADODB|DAO\.Recordset|rs\.Open)\b",
    re.IGNORECASE,
)
_BINARY_FIELDS = {FieldType.OLE_OBJECT, FieldType.ATTACHMENT}


@dataclass
class Recommendation:
    category: str
    priority: str  # "blocker", "high", "medium", "low", "info"
    title: str
    detail: str
    objects: list[str] = field(default_factory=list)


@dataclass
class AnalysisReport:
    catalog_file: str
    total_tables: int
    total_queries: int
    total_forms: int
    total_reports: int
    total_modules: int
    total_vba_lines: int
    effort_tier: str  # Small / Medium / Large / Enterprise
    recommendations: list[Recommendation] = field(default_factory=list)
    complexity_distribution: dict[str, int] = field(default_factory=dict)


def _effort_tier(catalog: AccessCatalog) -> str:
    total = (
        len(catalog.tables) + len(catalog.queries) +
        len(catalog.forms) + len(catalog.reports) + len(catalog.modules)
    )
    has_pass_through = any(q.is_pass_through for q in catalog.queries)
    has_active_x = any(
        any(c.is_active_x for c in f.controls) for f in catalog.forms
    )
    vba_locked = catalog.properties.vba_project_locked
    has_workgroup = bool(catalog.properties.workgroup_db_path)
    linked_external = [t for t in catalog.tables if t.is_linked and t.linked_source]

    if vba_locked or has_workgroup or len(linked_external) >= 3:
        return "Enterprise"
    if total > 200 or (has_pass_through and total > 50) or has_active_x:
        return "Large"
    if total > 50 or has_pass_through:
        return "Medium"
    return "Small"


def generate_recommendations(catalog: AccessCatalog) -> AnalysisReport:
    recs: list[Recommendation] = []

    # -----------------------------------------------------------------------
    # MIGRATION BLOCKERS
    # -----------------------------------------------------------------------

    if catalog.properties.vba_project_locked:
        recs.append(Recommendation(
            category="MIGRATION_BLOCKERS",
            priority="blocker",
            title="VBA project is password-protected",
            detail=(
                "The VBA source code cannot be extracted without the VBA project password. "
                "You must obtain the unlocked database file. This blocks all VBA migration planning."
            ),
        ))

    if catalog.properties.workgroup_db_path:
        recs.append(Recommendation(
            category="MIGRATION_BLOCKERS",
            priority="blocker",
            title="Workgroup (user-level) security detected",
            detail=(
                f"This database uses workgroup security (.mdw file: "
                f"{catalog.properties.workgroup_db_path}). "
                "Re-run with --workgroup-file to attempt full extraction. "
                "Workgroup security has no direct equivalent in modern databases."
            ),
        ))

    active_x_controls = []
    for form in catalog.forms:
        for ctrl in form.controls:
            if ctrl.is_active_x and ctrl.active_x_prog_id:
                active_x_controls.append(f"{form.name}/{ctrl.name} ({ctrl.active_x_prog_id})")

    if active_x_controls:
        unique_prog_ids = list({
            ctrl.active_x_prog_id
            for form in catalog.forms
            for ctrl in form.controls
            if ctrl.is_active_x and ctrl.active_x_prog_id
        })
        recs.append(Recommendation(
            category="MIGRATION_BLOCKERS",
            priority="blocker",
            title=f"ActiveX controls ({len(active_x_controls)} instances, {len(unique_prog_ids)} unique ProgIDs)",
            detail=(
                "ActiveX controls have no direct web/desktop framework equivalent. "
                "Each ProgID must be evaluated individually for a replacement component. "
                f"Unique ProgIDs: {unique_prog_ids[:8]}"
            ),
            objects=active_x_controls[:10],
        ))

    linked_tables = [t for t in catalog.tables if t.is_linked and t.linked_source]
    if linked_tables:
        sources = list({
            t.linked_source.connect_string[:40]
            for t in linked_tables if t.linked_source
        })
        recs.append(Recommendation(
            category="MIGRATION_BLOCKERS",
            priority="high",
            title=f"Linked tables from external sources ({len(linked_tables)} tables)",
            detail=(
                "Linked tables point to external data sources that must be migrated separately. "
                f"Connection string hints: {sources[:4]}"
            ),
            objects=[t.name for t in linked_tables],
        ))

    broken_refs = [r for r in catalog.properties.vba_references if r.is_broken]
    if broken_refs:
        recs.append(Recommendation(
            category="MIGRATION_BLOCKERS",
            priority="high",
            title=f"Broken VBA references ({len(broken_refs)})",
            detail=(
                "VBA code references libraries that are not installed on this machine. "
                "The code may be partially non-functional. Resolve before migration. "
                f"Broken refs: {[r.name for r in broken_refs]}"
            ),
            objects=[r.name for r in broken_refs],
        ))

    # -----------------------------------------------------------------------
    # SCHEMA_ISSUES
    # -----------------------------------------------------------------------

    no_pk_tables = [
        t for t in catalog.tables
        if not t.is_system_table and not t.is_linked
        and not any(f.is_primary_key for f in t.fields)
        and not any(idx.primary for idx in t.indexes)
        and t.fields
    ]
    if no_pk_tables:
        recs.append(Recommendation(
            category="SCHEMA_ISSUES",
            priority="high",
            title=f"Tables without primary keys ({len(no_pk_tables)})",
            detail="Tables without PKs complicate referential integrity and ORM mapping.",
            objects=[t.name for t in no_pk_tables],
        ))

    binary_tables = {}
    for t in catalog.tables:
        binary_flds = [f.name for f in t.fields if f.field_type in _BINARY_FIELDS]
        if binary_flds:
            binary_tables[t.name] = binary_flds
    if binary_tables:
        recs.append(Recommendation(
            category="SCHEMA_ISSUES",
            priority="high",
            title=f"OLE Object / Attachment fields ({len(binary_tables)} tables)",
            detail=(
                "Binary fields cannot be migrated automatically. "
                "Define a storage strategy: filesystem paths, S3/blob references, or DB blobs. "
                "Each field requires a manual migration plan."
            ),
            objects=[f"{t}: {flds}" for t, flds in list(binary_tables.items())[:8]],
        ))

    unenforced_rels = [r for r in catalog.relations if not r.enforce_integrity]
    if unenforced_rels:
        recs.append(Recommendation(
            category="SCHEMA_ISSUES",
            priority="medium",
            title=f"Relationships without referential integrity ({len(unenforced_rels)})",
            detail=(
                "Consider adding FK constraints in the target database. "
                "Verify that the actual data is consistent before enforcing."
            ),
            objects=[f"{r.parent_table} → {r.child_table}" for r in unenforced_rels[:10]],
        ))

    space_name_tables = [
        t for t in catalog.tables
        if any(" " in f.name for f in t.fields)
    ]
    if space_name_tables:
        recs.append(Recommendation(
            category="SCHEMA_ISSUES",
            priority="low",
            title=f"Field names with spaces ({len(space_name_tables)} tables)",
            detail=(
                "Access allows spaces in field names, but most databases and ORMs don't. "
                "Map to snake_case in the target schema."
            ),
            objects=[t.name for t in space_name_tables],
        ))

    # -----------------------------------------------------------------------
    # QUERY_COMPLEXITY
    # -----------------------------------------------------------------------

    pass_through = [q for q in catalog.queries if q.is_pass_through]
    if pass_through:
        recs.append(Recommendation(
            category="QUERY_COMPLEXITY",
            priority="high",
            title=f"Pass-through queries ({len(pass_through)})",
            detail=(
                "Pass-through queries contain server-side SQL (often T-SQL) "
                "that won't run against the new target database without modification."
            ),
            objects=[q.name for q in pass_through],
        ))

    queries_with_domain_agg = [
        q for q in catalog.queries
        if _DOMAIN_AGGREGATE_RE.search(q.sql_text)
    ]
    if queries_with_domain_agg:
        recs.append(Recommendation(
            category="QUERY_COMPLEXITY",
            priority="high",
            title=f"Domain aggregate functions ({len(queries_with_domain_agg)} queries)",
            detail=(
                "DLookup, DSum, DCount, etc. have no direct SQL equivalent. "
                "Replace with correlated subqueries, CTEs, or application-layer logic."
            ),
            objects=[q.name for q in queries_with_domain_agg[:10]],
        ))

    circular_queries = [q for q in catalog.queries if q.in_circular_dependency]
    if circular_queries:
        recs.append(Recommendation(
            category="QUERY_COMPLEXITY",
            priority="high",
            title=f"Circular query dependencies ({len(circular_queries)} queries)",
            detail=(
                "These queries reference each other in a cycle. "
                "This is invalid in most database systems. "
                "Break the cycle before migration."
            ),
            objects=[q.name for q in circular_queries],
        ))

    action_queries = [
        q for q in catalog.queries
        if q.query_type in (QueryType.DELETE, QueryType.UPDATE, QueryType.APPEND, QueryType.MAKE_TABLE)
    ]
    if action_queries:
        recs.append(Recommendation(
            category="QUERY_COMPLEXITY",
            priority="medium",
            title=f"Action queries as workflows ({len(action_queries)})",
            detail=(
                "Action queries (delete, update, append, make-table) may represent business workflows. "
                "Consider replacing with stored procedures or application-layer transactions."
            ),
            objects=[f"{q.name} ({q.query_type.value})" for q in action_queries[:10]],
        ))

    # -----------------------------------------------------------------------
    # VBA_MIGRATION
    # -----------------------------------------------------------------------

    total_vba_lines = catalog.total_vba_lines()
    if total_vba_lines > 0:
        recs.append(Recommendation(
            category="VBA_MIGRATION",
            priority="info",
            title=f"Total VBA code: {total_vba_lines:,} lines across {len(catalog.modules)} modules",
            detail="Each module will need to be analyzed and rewritten in the target language.",
        ))

    # Top 10 most complex modules by LOC
    large_modules = sorted(
        [m for m in catalog.modules if m.line_count and m.line_count > 50],
        key=lambda m: m.line_count or 0,
        reverse=True,
    )[:10]
    if large_modules:
        recs.append(Recommendation(
            category="VBA_MIGRATION",
            priority="medium",
            title=f"Top {len(large_modules)} largest VBA modules",
            detail="Prioritize reviewing these modules for business logic.",
            objects=[f"{m.name} ({m.line_count} lines)" for m in large_modules],
        ))

    # External COM dependencies
    non_builtin_refs = [
        r for r in catalog.properties.vba_references
        if not r.is_broken and r.name.lower() not in {"vba", "access", "dao"}
    ]
    if non_builtin_refs:
        recs.append(Recommendation(
            category="VBA_MIGRATION",
            priority="medium",
            title=f"External VBA library dependencies ({len(non_builtin_refs)})",
            detail=(
                "VBA code depends on these COM libraries. "
                "Each must be replaced with an equivalent in the target stack."
            ),
            objects=[f"{r.name}: {r.description}" for r in non_builtin_refs[:10]],
        ))

    # On Error Resume Next
    oern_modules = [
        m for m in catalog.modules
        if m.source_code and _ON_ERROR_RESUME_RE.search(m.source_code)
    ]
    if oern_modules:
        recs.append(Recommendation(
            category="VBA_MIGRATION",
            priority="medium",
            title=f"'On Error Resume Next' usage ({len(oern_modules)} modules)",
            detail=(
                "This pattern silently suppresses errors. "
                "The code may have hidden bugs that surface during migration. "
                "Review carefully."
            ),
            objects=[m.name for m in oern_modules],
        ))

    # DAO/ADO patterns
    dao_modules = [
        m for m in catalog.modules
        if m.source_code and _DAO_ADO_RE.search(m.source_code)
    ]
    if dao_modules:
        recs.append(Recommendation(
            category="VBA_MIGRATION",
            priority="medium",
            title=f"Direct database access in VBA ({len(dao_modules)} modules)",
            detail=(
                "DAO/ADO patterns (CurrentDb, OpenRecordset, ADODB) indicate "
                "SQL strings built in code — possible SQL injection risk. "
                "Replace with parameterized queries or ORM calls."
            ),
            objects=[m.name for m in dao_modules[:10]],
        ))

    # -----------------------------------------------------------------------
    # DATA_QUALITY (scope=full only)
    # -----------------------------------------------------------------------

    if catalog.scope == ExtractionScope.FULL and catalog.table_profiles:
        from access_dissect.catalog.models import FieldProfile  # noqa

        high_null_fields = []
        type_mismatch_fields = []
        constant_fields = []

        for tp in catalog.table_profiles:
            for fp in tp.field_profiles:
                if fp.null_rate is not None and fp.null_rate > 0.8 and (fp.non_null_count or 0) > 0:
                    high_null_fields.append(f"{tp.table_name}.{fp.field_name} ({fp.null_rate*100:.0f}% NULL)")
                if fp.suspected_type:
                    type_mismatch_fields.append(
                        f"{tp.table_name}.{fp.field_name} ({fp.suspected_type})"
                    )
                if fp.is_constant and (fp.non_null_count or 0) > 1:
                    constant_fields.append(f"{tp.table_name}.{fp.field_name}")

        if high_null_fields:
            recs.append(Recommendation(
                category="DATA_QUALITY",
                priority="low",
                title=f"High-null fields ({len(high_null_fields)})",
                detail="Fields with >80% NULL values. Consider whether they are still needed.",
                objects=high_null_fields[:10],
            ))
        if type_mismatch_fields:
            recs.append(Recommendation(
                category="DATA_QUALITY",
                priority="medium",
                title=f"Suspected type mismatches ({len(type_mismatch_fields)} fields)",
                detail="Fields that may store dates or numbers as text. Clean up before migration.",
                objects=type_mismatch_fields[:10],
            ))
        if constant_fields:
            recs.append(Recommendation(
                category="DATA_QUALITY",
                priority="info",
                title=f"Constant-value fields ({len(constant_fields)})",
                detail="These fields contain only one distinct value. Consider a DEFAULT or removing them.",
                objects=constant_fields[:10],
            ))

    # -----------------------------------------------------------------------
    # QUICK_WINS
    # -----------------------------------------------------------------------

    clean_tables = [
        t for t in catalog.tables
        if not t.is_linked and not t.is_system_table
        and any(f.is_primary_key for f in t.fields)
        and not any(f.field_type in _BINARY_FIELDS for f in t.fields)
        and not any(" " in f.name for f in t.fields)
    ]
    if clean_tables:
        recs.append(Recommendation(
            category="QUICK_WINS",
            priority="info",
            title=f"Clean tables ({len(clean_tables)}) — straightforward migration candidates",
            detail="These tables have PKs, no binary fields, and clean naming.",
            objects=[t.name for t in clean_tables[:10]],
        ))

    simple_queries = [
        q for q in catalog.queries
        if q.query_type == QueryType.SELECT
        and not q.is_pass_through
        and not q.in_circular_dependency
        and not _DOMAIN_AGGREGATE_RE.search(q.sql_text)
        and not q.parameters
    ]
    if simple_queries:
        recs.append(Recommendation(
            category="QUICK_WINS",
            priority="info",
            title=f"Simple SELECT queries ({len(simple_queries)}) — may translate directly",
            detail=(
                "These queries use standard SQL without Access-specific functions. "
                "sqlglot can translate them automatically."
            ),
            objects=[q.name for q in simple_queries[:10]],
        ))

    effort = _effort_tier(catalog)
    from access_dissect.analyze.complexity import score_all_objects  # noqa

    scores = score_all_objects(catalog)
    dist = {"low": 0, "medium": 0, "high": 0}
    for s in scores:
        dist[s.band] += 1

    return AnalysisReport(
        catalog_file=catalog.properties.file_path,
        total_tables=len(catalog.tables),
        total_queries=len(catalog.queries),
        total_forms=len(catalog.forms),
        total_reports=len(catalog.reports),
        total_modules=len(catalog.modules),
        total_vba_lines=catalog.total_vba_lines(),
        effort_tier=effort,
        recommendations=recs,
        complexity_distribution=dist,
    )
