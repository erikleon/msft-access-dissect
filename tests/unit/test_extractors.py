"""
Unit tests for extractors using mock COM objects.

All tests run without Microsoft Access installed.
"""

from __future__ import annotations

from unittest.mock import MagicMock, patch

import pytest

from access_dissect.catalog.models import FieldType, QueryType
from access_dissect.extract.tables import extract_table, extract_all_tables
from access_dissect.extract.queries import extract_query, extract_all_queries
from access_dissect.extract.relations import extract_relation, extract_all_relations


class TestTableExtractor:
    def test_extract_simple_table(self, mock_tabledef_simple: MagicMock) -> None:
        table, warnings = extract_table(mock_tabledef_simple)
        assert table is not None
        assert table.name == "Customers"
        assert len(table.fields) == 3
        assert table.is_linked is False
        assert table.is_system_table is False

    def test_field_types_mapped(self, mock_tabledef_simple: MagicMock) -> None:
        table, _ = extract_table(mock_tabledef_simple)
        assert table is not None
        # CustomerID (type=4) → LONG
        customer_id = table.fields[0]
        assert customer_id.field_type == FieldType.LONG
        # CompanyName (type=10) → TEXT
        company_name = table.fields[1]
        assert company_name.field_type == FieldType.TEXT

    def test_primary_key_detected(self, mock_tabledef_simple: MagicMock) -> None:
        table, _ = extract_table(mock_tabledef_simple)
        assert table is not None
        customer_id = table.fields[0]
        assert customer_id.is_primary_key is True

    def test_msys_table_skipped_by_default(self) -> None:
        td = MagicMock()
        td.Name = "MSysObjects"
        td.Attributes = 0
        td.Connect = ""
        td.Fields = MagicMock()
        td.Fields.Count = 0
        td.Fields.__iter__ = lambda s: iter([])
        td.Indexes = MagicMock()
        td.Indexes.Count = 0
        td.Indexes.__iter__ = lambda s: iter([])
        table, _ = extract_table(td, include_system=False)
        assert table is None

    def test_msys_table_included_when_requested(self) -> None:
        td = MagicMock()
        td.Name = "MSysObjects"
        td.Attributes = 2147483650  # DB_SYSTEMOBJECT
        td.Connect = ""
        td.SourceTableName = ""
        td.Description = None
        fields_mock = MagicMock()
        fields_mock.Count = 0
        fields_mock.__iter__ = lambda s: iter([])
        td.Fields = fields_mock
        indexes_mock = MagicMock()
        indexes_mock.Count = 0
        indexes_mock.__iter__ = lambda s: iter([])
        td.Indexes = indexes_mock
        table, _ = extract_table(td, include_system=True)
        assert table is not None
        assert table.name == "MSysObjects"

    def test_linked_table_detected(self) -> None:
        td = MagicMock()
        td.Name = "LinkedCustomers"
        td.Attributes = 0x20000000  # DB_ATTACHED_ODBC
        td.Connect = "ODBC;DSN=Production;UID=user;PWD=secret;"
        td.SourceTableName = "dbo.Customers"
        td.Description = None

        fields_mock = MagicMock()
        fields_mock.Count = 0
        fields_mock.__iter__ = lambda s: iter([])
        td.Fields = fields_mock

        indexes_mock = MagicMock()
        indexes_mock.Count = 0
        indexes_mock.__iter__ = lambda s: iter([])
        td.Indexes = indexes_mock

        table, warnings = extract_table(td)
        assert table is not None
        assert table.is_linked is True
        assert table.linked_source is not None
        # Credentials should be stripped
        assert "PWD=secret" not in table.linked_source.connect_string

    def test_no_pk_generates_warning(self) -> None:
        td = MagicMock()
        td.Name = "NoPKTable"
        td.Attributes = 0
        td.Connect = ""
        td.SourceTableName = ""
        td.Description = None

        from tests.conftest import _mock_field, _iterable_mock
        fields_mock = _iterable_mock([_mock_field("Name", 0)])
        td.Fields = fields_mock
        td.Indexes = _iterable_mock([])

        table, warnings = extract_table(td)
        assert table is not None
        assert any("primary key" in w.message.lower() for w in warnings)

    def test_extract_all_tables_isolates_errors(self, mock_database: MagicMock) -> None:
        tables, warnings = extract_all_tables(mock_database)
        assert isinstance(tables, list)
        assert isinstance(warnings, list)


class TestQueryExtractor:
    def test_extract_simple_query(self, mock_querydef_simple: MagicMock) -> None:
        query, warnings = extract_query(mock_querydef_simple)
        assert query is not None
        assert query.name == "qryCustomers"
        assert query.query_type == QueryType.SELECT
        assert "Customers" in query.sql_text

    def test_sql_stored_verbatim(self, mock_querydef_simple: MagicMock) -> None:
        """SQL must not be modified during extraction."""
        original_sql = mock_querydef_simple.SQL
        query, _ = extract_query(mock_querydef_simple)
        assert query is not None
        assert query.sql_text == original_sql

    def test_tilde_query_skipped(self) -> None:
        qd = MagicMock()
        qd.Name = "~TempQuery"
        query, _ = extract_query(qd)
        assert query is None

    def test_pass_through_detected(self) -> None:
        qd = MagicMock()
        qd.Name = "qPassThru"
        qd.Type = 7  # dbQPassThrough
        qd.SQL = "SELECT * FROM dbo.Customers"
        qd.Parameters = MagicMock()
        qd.Parameters.Count = 0
        qd.Parameters.__iter__ = lambda s: iter([])
        qd.ReturnsRecords = True
        qd.Connect = "ODBC;DSN=Prod;"

        query, warnings = extract_query(qd)
        assert query is not None
        assert query.is_pass_through is True
        assert any("pass-through" in w.message.lower() for w in warnings)


class TestRelationExtractor:
    def test_extract_relation(self) -> None:
        rel = MagicMock()
        rel.Name = "OrdersCustomers"
        rel.Table = "Customers"
        rel.ForeignTable = "Orders"
        rel.Attributes = 256  # cascade update, RI enforced

        field1 = MagicMock()
        field1.Name = "CustomerID"
        field1.ForeignName = "CustomerID"
        fields_mock = MagicMock()
        fields_mock.__iter__ = lambda s: iter([field1])
        rel.Fields = fields_mock

        relation, warnings = extract_relation(rel)
        assert relation is not None
        assert relation.parent_table == "Customers"
        assert relation.child_table == "Orders"
        assert len(relation.fields) == 1
        assert relation.cascade_update is True

    def test_msys_relation_skipped(self, mock_database: MagicMock) -> None:
        """MSysRelationships system relations should be filtered out."""
        msys_rel = MagicMock()
        msys_rel.Name = "MSysRelationship1"
        msys_rel.Table = "MSysObjects"
        msys_rel.ForeignTable = "MSysQueries"
        msys_rel.Attributes = 2

        fields_mock = MagicMock()
        fields_mock.__iter__ = lambda s: iter([])
        msys_rel.Fields = fields_mock

        rels_mock = MagicMock()
        rels_mock.Count = 1
        rels_mock.__call__ = lambda i: msys_rel

        mock_database.Relations = rels_mock
        relations, _ = extract_all_relations(mock_database)
        # MSys relation should be filtered
        assert all(not r.name.startswith("MSys") for r in relations)
