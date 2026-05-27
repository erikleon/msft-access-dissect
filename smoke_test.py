"""Quick end-to-end smoke test of the full pipeline (no Access required)."""
import tempfile
from datetime import datetime
from pathlib import Path

from access_dissect.catalog.models import (
    AccessCatalog, DatabaseProperties, ExtractionScope,
    FieldDef, FieldType, FormDef, IndexDef, IndexField,
    QueryDef, QueryType, RelationDef, RelationField,
    TableDef, VBAModule, VBAModuleType, VBAReference,
)
from access_dissect.utils.io import save_catalog, load_catalog
from access_dissect.render.markdown import render_markdown
from access_dissect.render.html import render_html
from access_dissect.render.sql import generate_ddl, save_ddl
from access_dissect.analyze.engine import analyze_catalog, render_analysis_markdown

catalog = AccessCatalog(
    extracted_at=datetime(2026, 1, 15),
    extracted_by="access-dissect 0.1.0",
    scope=ExtractionScope.STRUCTURE,
    properties=DatabaseProperties(
        title="Northwind",
        file_path=r"C:\Northwind.accdb",
        is_accdb=True,
        access_version="16.0",
        vba_references=[
            VBAReference(name="VBA", description="VBA"),
            VBAReference(name="DAO", description="DAO 3.6"),
        ],
    ),
    tables=[
        TableDef(
            name="Customers",
            fields=[
                FieldDef(name="CustomerID", ordinal=0, field_type=FieldType.LONG,
                         dao_type_code=4, is_primary_key=True, auto_number=True),
                FieldDef(name="CompanyName", ordinal=1, field_type=FieldType.TEXT,
                         dao_type_code=10, size=50, required=True),
            ],
            indexes=[IndexDef(
                name="PrimaryKey",
                fields=[IndexField(field_name="CustomerID")],
                primary=True, unique=True,
            )],
        ),
        TableDef(
            name="Orders",
            fields=[
                FieldDef(name="OrderID", ordinal=0, field_type=FieldType.LONG,
                         dao_type_code=4, is_primary_key=True, auto_number=True),
                FieldDef(name="CustomerID", ordinal=1, field_type=FieldType.LONG,
                         dao_type_code=4, required=True),
            ],
        ),
    ],
    queries=[
        QueryDef(name="qryAll", query_type=QueryType.SELECT, sql_text="SELECT * FROM Customers;"),
        QueryDef(name="qCircA", query_type=QueryType.SELECT, sql_text="SELECT * FROM qCircB"),
        QueryDef(name="qCircB", query_type=QueryType.SELECT, sql_text="SELECT * FROM qCircA"),
    ],
    relations=[
        RelationDef(
            name="FK_Orders",
            parent_table="Customers",
            child_table="Orders",
            fields=[RelationField(parent_field="CustomerID", child_field="CustomerID")],
            enforce_integrity=True,
        )
    ],
    modules=[
        VBAModule(
            name="modUtils",
            module_type=VBAModuleType.STANDARD,
            source_code='Function Hello() As String\n    Hello = "World"\nEnd Function',
            line_count=3,
        )
    ],
    forms=[
        FormDef(name="frmCustomers", record_source="Customers", has_vba_module=False),
    ],
)

with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)

    # JSON roundtrip
    cat_path = tmp / "catalog.json"
    save_catalog(catalog, cat_path)
    loaded = load_catalog(cat_path)
    assert len(loaded.tables) == 2
    assert loaded.properties.title == "Northwind"
    print(f"✓ JSON roundtrip: {len(loaded.tables)} tables, {len(loaded.queries)} queries")

    # YAML roundtrip
    yaml_path = tmp / "catalog.yaml"
    save_catalog(catalog, yaml_path, fmt="yaml")
    loaded_yaml = load_catalog(yaml_path)
    assert loaded_yaml.scope == ExtractionScope.STRUCTURE
    print(f"✓ YAML roundtrip: {len(loaded_yaml.modules)} modules")

    # Markdown
    md_dir = tmp / "docs"
    render_markdown(catalog, md_dir)
    md_files = list(md_dir.rglob("*.md"))
    assert (md_dir / "README.md").exists()
    assert (md_dir / "tables" / "Customers.md").exists()
    assert (md_dir / "relations.md").exists()
    readme = (md_dir / "README.md").read_text()
    assert "Northwind" in readme
    print(f"✓ Markdown: {len(md_files)} files")

    # HTML
    html_path = tmp / "report.html"
    render_html(catalog, html_path)
    html_size = html_path.stat().st_size
    assert html_size > 5000
    html_content = html_path.read_text(encoding="utf-8")
    assert "Northwind" in html_content
    assert "erDiagram" in html_content
    print(f"✓ HTML report: {html_size:,} bytes")

    # SQL DDL (all 3 dialects)
    for dialect in ("postgres", "sqlite", "mssql"):
        sql_path = tmp / f"schema_{dialect}.sql"
        save_ddl(catalog, sql_path, dialect=dialect)
        ddl = sql_path.read_text()
        assert "CREATE TABLE" in ddl
        assert "FOREIGN KEY" in ddl
        print(f"✓ SQL DDL ({dialect}): {len(ddl)} chars")

    # Analysis
    report = analyze_catalog(catalog, verbose=False)
    assert report.effort_tier in ("Small", "Medium", "Large", "Enterprise")
    assert len(report.recommendations) > 0
    analysis_path = tmp / "analysis.md"
    render_analysis_markdown(report, analysis_path)
    analysis_text = analysis_path.read_text(encoding="utf-8")
    assert "Complexity" in analysis_text or "Migration" in analysis_text or "Circular" in analysis_text
    print(f"✓ Analysis: tier={report.effort_tier}, {len(report.recommendations)} recs, {len(analysis_text)} chars")

    # Verify circular queries flagged
    assert catalog.queries[1].in_circular_dependency or catalog.queries[2].in_circular_dependency
    print("✓ Circular query dependencies detected")

    # Verify dependency graph populated
    assert len(catalog.dependency_graph) > 0
    print(f"✓ Dependency graph: {len(catalog.dependency_graph)} nodes")

print()
print("=" * 50)
print("ALL SMOKE TESTS PASSED ✓")
print("=" * 50)
