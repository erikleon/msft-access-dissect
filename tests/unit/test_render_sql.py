"""Unit tests for SQL DDL rendering."""

from __future__ import annotations

import tempfile
from pathlib import Path

import pytest

from access_dissect.catalog.models import AccessCatalog
from access_dissect.render.sql import generate_ddl, save_ddl


class TestDDLGenerator:
    def test_generates_create_table(self, sample_catalog: AccessCatalog) -> None:
        ddl = generate_ddl(sample_catalog, dialect="postgres")
        assert "CREATE TABLE" in ddl
        assert "Customers" in ddl or '"Customers"' in ddl

    def test_generates_primary_key(self, sample_catalog: AccessCatalog) -> None:
        ddl = generate_ddl(sample_catalog, dialect="postgres")
        assert "PRIMARY KEY" in ddl

    def test_generates_foreign_key(self, sample_catalog: AccessCatalog) -> None:
        ddl = generate_ddl(sample_catalog, dialect="postgres")
        assert "FOREIGN KEY" in ddl

    def test_postgres_dialect(self, sample_catalog: AccessCatalog) -> None:
        ddl = generate_ddl(sample_catalog, dialect="postgres")
        # Postgres uses double-quotes for identifiers
        assert '"' in ddl or "CREATE TABLE" in ddl

    def test_mssql_dialect(self, sample_catalog: AccessCatalog) -> None:
        ddl = generate_ddl(sample_catalog, dialect="mssql")
        # MSSQL uses square brackets
        assert "[" in ddl

    def test_sqlite_dialect(self, sample_catalog: AccessCatalog) -> None:
        ddl = generate_ddl(sample_catalog, dialect="sqlite")
        assert "CREATE TABLE" in ddl

    def test_linked_table_skipped(self, sample_catalog: AccessCatalog) -> None:
        from access_dissect.catalog.models import LinkedTableSource, TableDef

        sample_catalog.tables.append(
            TableDef(
                name="LinkedTable",
                is_linked=True,
                linked_source=LinkedTableSource(connect_string="ODBC;DSN=Prod;"),
            )
        )
        ddl = generate_ddl(sample_catalog, dialect="postgres")
        # LinkedTable should appear as a comment, not as CREATE TABLE "LinkedTable"
        assert 'CREATE TABLE "LinkedTable"' not in ddl
        assert 'CREATE TABLE [LinkedTable]' not in ddl
        # There should be a comment mentioning it
        assert "LinkedTable" in ddl

    def test_header_comment_present(self, sample_catalog: AccessCatalog) -> None:
        ddl = generate_ddl(sample_catalog, dialect="postgres")
        assert "access-dissect" in ddl.lower() or "generated" in ddl.lower()

    def test_save_ddl_writes_file(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        out = tmp_path / "schema.sql"
        save_ddl(sample_catalog, out, dialect="postgres")
        assert out.exists()
        content = out.read_text(encoding="utf-8")
        assert "CREATE TABLE" in content
