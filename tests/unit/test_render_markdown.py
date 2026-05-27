"""Unit tests for the Markdown renderer output shape."""

from __future__ import annotations

import tempfile
from pathlib import Path

import pytest

from access_dissect.catalog.models import AccessCatalog
from access_dissect.render.markdown import render_markdown


class TestMarkdownRenderer:
    def test_creates_readme(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        render_markdown(sample_catalog, tmp_path)
        assert (tmp_path / "README.md").exists()

    def test_creates_table_files(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        render_markdown(sample_catalog, tmp_path)
        tables_dir = tmp_path / "tables"
        assert tables_dir.exists()
        assert (tables_dir / "Customers.md").exists()
        assert (tables_dir / "Orders.md").exists()

    def test_creates_query_files(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        render_markdown(sample_catalog, tmp_path)
        queries_dir = tmp_path / "queries"
        assert queries_dir.exists()
        assert (queries_dir / "qryCustomerOrders.md").exists()

    def test_creates_module_files(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        render_markdown(sample_catalog, tmp_path)
        mods_dir = tmp_path / "modules"
        assert mods_dir.exists()
        assert (mods_dir / "modUtilities.md").exists()

    def test_creates_relations_file(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        render_markdown(sample_catalog, tmp_path)
        assert (tmp_path / "relations.md").exists()

    def test_creates_warnings_file(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        render_markdown(sample_catalog, tmp_path)
        assert (tmp_path / "warnings.md").exists()

    def test_readme_contains_object_counts(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        render_markdown(sample_catalog, tmp_path)
        readme = (tmp_path / "README.md").read_text(encoding="utf-8")
        assert "Tables" in readme
        assert "Queries" in readme

    def test_table_doc_contains_fields(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        render_markdown(sample_catalog, tmp_path)
        content = (tmp_path / "tables" / "Customers.md").read_text(encoding="utf-8")
        assert "CustomerID" in content
        assert "CompanyName" in content

    def test_table_doc_contains_pk_marker(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        render_markdown(sample_catalog, tmp_path)
        content = (tmp_path / "tables" / "Customers.md").read_text(encoding="utf-8")
        assert "PK" in content or "🔑" in content

    def test_query_doc_contains_sql(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        render_markdown(sample_catalog, tmp_path)
        content = (tmp_path / "queries" / "qryCustomerOrders.md").read_text(encoding="utf-8")
        assert "SELECT" in content
        assert "Customers" in content

    def test_module_doc_contains_source(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        render_markdown(sample_catalog, tmp_path)
        content = (tmp_path / "modules" / "modUtilities.md").read_text(encoding="utf-8")
        assert "FormatPhone" in content

    def test_relations_doc_contains_mermaid(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        render_markdown(sample_catalog, tmp_path)
        content = (tmp_path / "relations.md").read_text(encoding="utf-8")
        assert "mermaid" in content
        assert "erDiagram" in content

    def test_warnings_doc_no_warnings(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        render_markdown(sample_catalog, tmp_path)
        content = (tmp_path / "warnings.md").read_text(encoding="utf-8")
        assert "No warnings" in content
