"""
MarkdownRenderer: generate one Markdown file per object type, cross-linked.

Output structure:
  output_dir/
    README.md              (index with summary and links)
    tables/
      <TableName>.md
    queries/
      <QueryName>.md
    forms/
      <FormName>.md
    reports/
      <ReportName>.md
    modules/
      <ModuleName>.md
    relations.md
    warnings.md
"""

from __future__ import annotations

import re
from pathlib import Path

from access_dissect.catalog.models import (
    AccessCatalog,
    FormDef,
    QueryDef,
    ReportDef,
    TableDef,
    VBAModule,
    WarningSeverity,
)
from access_dissect.translate.type_map import field_type_to_sql


def _safe_filename(name: str) -> str:
    return re.sub(r"[^\w\-]", "_", name) + ".md"


def _table_row(cells: list[str]) -> str:
    return "| " + " | ".join(str(c) for c in cells) + " |"


def _table_sep(n: int) -> str:
    return "| " + " | ".join(["---"] * n) + " |"


def _render_table_doc(table: TableDef, dialect: str = "postgres") -> str:
    lines: list[str] = []
    lines.append(f"# Table: {table.name}")
    lines.append("")

    if table.description:
        lines.append(f"_{table.description}_")
        lines.append("")

    # Metadata badges
    badges = []
    if table.is_linked:
        badges.append("🔗 Linked Table")
    if table.is_hidden:
        badges.append("👁 Hidden")
    if table.record_count is not None:
        lines.append(f"**Row count**: {table.record_count:,}")
    if badges:
        lines.append(" · ".join(badges))
    lines.append("")

    # Linked source info
    if table.is_linked and table.linked_source:
        ls = table.linked_source
        lines.append("## External Source")
        lines.append(f"- **Connection**: `{ls.connect_string}`")
        if ls.source_table_name:
            lines.append(f"- **Source table**: `{ls.source_table_name}`")
        lines.append("")

    # Fields
    lines.append("## Fields")
    lines.append("")
    lines.append(_table_row(["Name", "Type", f"SQL ({dialect})", "Size", "Required", "Default", "Notes"]))
    lines.append(_table_sep(7))
    for f in table.fields:
        sql_type = field_type_to_sql(f.field_type, dialect, f.size)
        notes = []
        if f.is_primary_key:
            notes.append("🔑 PK")
        if f.auto_number:
            notes.append("AutoNumber")
        if f.validation_rule:
            notes.append(f"Rule: `{f.validation_rule}`")
        lines.append(_table_row([
            f"`{f.name}`",
            f.field_type.value,
            sql_type,
            str(f.size) if f.size else "",
            "✓" if f.required else "",
            f"`{f.default_value}`" if f.default_value else "",
            " / ".join(notes),
        ]))
    lines.append("")

    # Indexes
    if table.indexes:
        lines.append("## Indexes")
        lines.append("")
        lines.append(_table_row(["Name", "Fields", "Unique", "Primary"]))
        lines.append(_table_sep(4))
        for idx in table.indexes:
            fields_str = ", ".join(
                f"`{f.field_name}`{'↓' if f.descending else ''}" for f in idx.fields
            )
            lines.append(_table_row([
                f"`{idx.name}`",
                fields_str,
                "✓" if idx.unique else "",
                "✓" if idx.primary else "",
            ]))
        lines.append("")

    return "\n".join(lines)


def _render_query_doc(query: QueryDef) -> str:
    lines: list[str] = []
    lines.append(f"# Query: {query.name}")
    lines.append("")

    badges = [f"Type: **{query.query_type.value}**"]
    if query.is_pass_through:
        badges.append("⚠ Pass-Through")
    if query.in_circular_dependency:
        badges.append("🔄 Circular Dependency")
    lines.append(" · ".join(badges))
    lines.append("")

    if query.parameters:
        lines.append("## Parameters")
        lines.append("")
        lines.append(_table_row(["Name", "Type"]))
        lines.append(_table_sep(2))
        for p in query.parameters:
            lines.append(_table_row([f"`{p.name}`", p.data_type]))
        lines.append("")

    if query.references_tables:
        lines.append("## References Tables")
        lines.append(", ".join(f"[{t}](../tables/{_safe_filename(t)})" for t in query.references_tables))
        lines.append("")

    if query.references_queries:
        lines.append("## References Queries")
        lines.append(", ".join(f"[{q}]({_safe_filename(q)})" for q in query.references_queries))
        lines.append("")

    lines.append("## SQL (Verbatim Jet/ACE)")
    lines.append("")
    lines.append("```sql")
    lines.append(query.sql_text)
    lines.append("```")
    lines.append("")

    return "\n".join(lines)


def _render_form_doc(form: FormDef) -> str:
    lines: list[str] = []
    lines.append(f"# Form: {form.name}")
    lines.append("")

    if form.record_source:
        lines.append(f"**Record Source**: `{form.record_source}`")
    if form.popup:
        lines.append("⚠ Modal / Popup form")
    lines.append("")

    # Controls summary
    if form.controls:
        lines.append(f"## Controls ({len(form.controls)})")
        lines.append("")
        lines.append(_table_row(["Name", "Type", "Bound Field", "Events"]))
        lines.append(_table_sep(4))
        for ctrl in form.controls[:50]:  # cap at 50 for readability
            event_summary = ", ".join(list(ctrl.events.keys())[:3]) if ctrl.events else ""
            lines.append(_table_row([
                f"`{ctrl.name}`",
                ctrl.control_type_name,
                f"`{ctrl.control_source}`" if ctrl.control_source else "",
                event_summary,
            ]))
        if len(form.controls) > 50:
            lines.append(f"\n_... {len(form.controls) - 50} more controls not shown_")
        lines.append("")

    # Subforms
    if form.subforms:
        lines.append("## Subforms")
        lines.append("")
        for sf in form.subforms:
            lines.append(f"- [{sf}](../forms/{_safe_filename(sf)})")
        lines.append("")

    # Events
    if form.events:
        lines.append("## Form Events")
        lines.append("")
        lines.append(_table_row(["Event", "Handler"]))
        lines.append(_table_sep(2))
        for evt, handler in form.events.items():
            lines.append(_table_row([evt, handler]))
        lines.append("")

    # VBA module link
    if form.has_vba_module and form.vba_module_name:
        lines.append(
            f"## VBA Code\n\nSee [module {form.vba_module_name}](../modules/{_safe_filename(form.vba_module_name)})"
        )
        lines.append("")

    # Extraction warnings
    if form.extraction_warnings:
        lines.append("## ⚠ Extraction Warnings")
        for w in form.extraction_warnings[:10]:
            lines.append(f"- {w}")
        lines.append("")

    return "\n".join(lines)


def _render_module_doc(module: VBAModule) -> str:
    lines: list[str] = []
    lines.append(f"# VBA Module: {module.name}")
    lines.append("")
    lines.append(f"**Type**: {module.module_type.value}")
    if module.parent_object:
        lines.append(f"**Parent**: {module.parent_object}")
    if module.line_count:
        lines.append(f"**Lines**: {module.line_count}")
    lines.append("")

    if module.is_vba_locked:
        lines.append(
            "> ⚠ **VBA project is password-protected.** "
            "Source code is not available. This is a migration blocker."
        )
        lines.append("")
    elif module.source_code:
        lines.append("## Source Code")
        lines.append("")
        lines.append("```vba")
        lines.append(module.source_code)
        lines.append("```")
        lines.append("")
    else:
        lines.append("_No source code extracted._")
        lines.append("")

    return "\n".join(lines)


def _render_index(catalog: AccessCatalog) -> str:
    lines: list[str] = []
    lines.append(f"# Access Database: {catalog.properties.title or 'Untitled'}")
    lines.append("")
    lines.append(f"> Generated by {catalog.extracted_by} on {catalog.extracted_at.strftime('%Y-%m-%d %H:%M')}")
    lines.append(f"> Source: `{catalog.properties.file_path}`")
    lines.append("")

    lines.append("## Summary")
    lines.append("")
    lines.append("| Object | Count |")
    lines.append("|--------|-------|")
    lines.append(f"| Tables | {len(catalog.tables)} |")
    lines.append(f"| Queries | {len(catalog.queries)} |")
    lines.append(f"| Forms | {len(catalog.forms)} |")
    lines.append(f"| Reports | {len(catalog.reports)} |")
    lines.append(f"| VBA Modules | {len(catalog.modules)} |")
    lines.append(f"| Macros | {len(catalog.macros)} |")
    lines.append(f"| Relationships | {len(catalog.relations)} |")
    lines.append(f"| Total VBA lines | {catalog.total_vba_lines():,} |")
    lines.append("")

    error_count = catalog.error_count()
    warn_count = len(catalog.warnings_by_severity(WarningSeverity.WARNING))
    if error_count:
        lines.append(f"> ⚠ **{error_count} extraction error(s)** — see [warnings](warnings.md)")
    elif warn_count:
        lines.append(f"> ℹ {warn_count} warning(s) — see [warnings](warnings.md)")
    lines.append("")

    for section, items, subdir in [
        ("Tables", catalog.tables, "tables"),
        ("Queries", catalog.queries, "queries"),
        ("Forms", catalog.forms, "forms"),
        ("Reports", catalog.reports, "reports"),
        ("VBA Modules", catalog.modules, "modules"),
    ]:
        if items:
            lines.append(f"## {section}")
            lines.append("")
            for item in items:
                lines.append(f"- [{item.name}]({subdir}/{_safe_filename(item.name)})")
            lines.append("")

    if catalog.relations:
        lines.append("## Relationships")
        lines.append("")
        lines.append("[View all relationships](relations.md)")
        lines.append("")

    return "\n".join(lines)


def _render_relations_doc(catalog: AccessCatalog) -> str:
    lines: list[str] = []
    lines.append("# Relationships")
    lines.append("")
    lines.append(_table_row(["Name", "Parent Table", "Child Table", "RI", "Cascade Update", "Cascade Delete", "1:1"]))
    lines.append(_table_sep(7))
    for rel in catalog.relations:
        lines.append(_table_row([
            rel.name,
            f"[{rel.parent_table}](tables/{_safe_filename(rel.parent_table)})",
            f"[{rel.child_table}](tables/{_safe_filename(rel.child_table)})",
            "✓" if rel.enforce_integrity else "",
            "✓" if rel.cascade_update else "",
            "✓" if rel.cascade_delete else "",
            "✓" if rel.one_to_one else "",
        ]))
    lines.append("")

    # ERD in Mermaid
    lines.append("## Entity Relationship Diagram")
    lines.append("")
    lines.append("```mermaid")
    lines.append("erDiagram")
    for rel in catalog.relations:
        lhs = f"{rel.parent_table}"
        rhs = f"{rel.child_table}"
        if rel.one_to_one:
            cardinality = "||--||"
        else:
            cardinality = "||--o{"
        label = "enforces" if rel.enforce_integrity else "references"
        lines.append(f'  {lhs} {cardinality} {rhs} : "{label}"')
    lines.append("```")
    lines.append("")

    return "\n".join(lines)


def _render_warnings_doc(catalog: AccessCatalog) -> str:
    lines: list[str] = []
    lines.append("# Extraction Warnings")
    lines.append("")

    if not catalog.warnings:
        lines.append("✓ No warnings.")
        return "\n".join(lines)

    for severity in (WarningSeverity.ERROR, WarningSeverity.WARNING, WarningSeverity.INFO):
        severity_warnings = [w for w in catalog.warnings if w.severity == severity]
        if not severity_warnings:
            continue
        lines.append(f"## {severity.value.upper()} ({len(severity_warnings)})")
        lines.append("")
        lines.append(_table_row(["Object", "Code", "Message"]))
        lines.append(_table_sep(3))
        for w in severity_warnings:
            lines.append(_table_row([
                f"{w.object_type}:{w.object_name}",
                f"`{w.warning_code}`",
                w.message[:120],
            ]))
        lines.append("")

    return "\n".join(lines)


def render_markdown(
    catalog: AccessCatalog,
    output_dir: Path,
    dialect: str = "postgres",
) -> None:
    """
    Render the full catalog as Markdown files into output_dir.
    Creates subdirectories for each object type.
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # README
    (output_dir / "README.md").write_text(_render_index(catalog), encoding="utf-8")

    # Tables
    tbl_dir = output_dir / "tables"
    tbl_dir.mkdir(parents=True, exist_ok=True)
    for table in catalog.tables:
        (tbl_dir / _safe_filename(table.name)).write_text(
            _render_table_doc(table, dialect), encoding="utf-8"
        )

    # Queries
    qry_dir = output_dir / "queries"
    qry_dir.mkdir(parents=True, exist_ok=True)
    for query in catalog.queries:
        (qry_dir / _safe_filename(query.name)).write_text(
            _render_query_doc(query), encoding="utf-8"
        )

    # Forms
    frm_dir = output_dir / "forms"
    frm_dir.mkdir(parents=True, exist_ok=True)
    for form in catalog.forms:
        (frm_dir / _safe_filename(form.name)).write_text(
            _render_form_doc(form), encoding="utf-8"
        )

    # Reports
    rpt_dir = output_dir / "reports"
    rpt_dir.mkdir(parents=True, exist_ok=True)
    for report in catalog.reports:
        (rpt_dir / _safe_filename(report.name)).write_text(
            _report_summary(report), encoding="utf-8"
        )

    # VBA modules
    mod_dir = output_dir / "modules"
    mod_dir.mkdir(parents=True, exist_ok=True)
    for module in catalog.modules:
        (mod_dir / _safe_filename(module.name)).write_text(
            _render_module_doc(module), encoding="utf-8"
        )

    # Relations
    (output_dir / "relations.md").write_text(
        _render_relations_doc(catalog), encoding="utf-8"
    )

    # Warnings
    (output_dir / "warnings.md").write_text(
        _render_warnings_doc(catalog), encoding="utf-8"
    )


def _report_summary(report: ReportDef) -> str:
    lines: list[str] = []
    lines.append(f"# Report: {report.name}")
    lines.append("")
    if report.record_source:
        lines.append(f"**Record Source**: `{report.record_source}`")
    if report.order_by:
        lines.append(f"**Order By**: `{report.order_by}`")
    if report.filter:
        lines.append(f"**Filter**: `{report.filter}`")
    lines.append("")

    if report.sections:
        lines.append("## Sections")
        for sec in report.sections:
            ctrl_count = len(sec.controls)
            lines.append(f"- **{sec.section_name}**: {ctrl_count} controls")
        lines.append("")

    if report.groupings:
        lines.append("## Groupings")
        for g in report.groupings:
            lines.append(f"- `{g.field_or_expression}`")
        lines.append("")

    if report.has_vba_module and report.vba_module_name:
        lines.append(
            f"## VBA Code\n\nSee [module {report.vba_module_name}](../modules/{_safe_filename(report.vba_module_name)})"
        )

    if report.extraction_warnings:
        lines.append("## ⚠ Extraction Warnings")
        for w in report.extraction_warnings[:10]:
            lines.append(f"- {w}")

    return "\n".join(lines)
