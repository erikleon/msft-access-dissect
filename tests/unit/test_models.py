"""Unit tests for catalog Pydantic models."""

from __future__ import annotations

import json
from datetime import datetime

import pytest

from access_dissect.catalog.models import (
    AccessCatalog,
    DatabaseProperties,
    ExtractionScope,
    ExtractionWarning,
    FieldDef,
    FieldType,
    FormDef,
    QueryDef,
    QueryType,
    RelationDef,
    RelationField,
    TableDef,
    VBAModule,
    VBAModuleType,
    WarningSeverity,
)


def make_minimal_catalog() -> AccessCatalog:
    return AccessCatalog(
        extracted_at=datetime(2026, 1, 1),
        extracted_by="test",
        scope=ExtractionScope.STRUCTURE,
        properties=DatabaseProperties(file_path="C:\\test.accdb"),
    )


class TestAccessCatalogRoundtrip:
    def test_json_roundtrip_empty(self) -> None:
        catalog = make_minimal_catalog()
        json_str = catalog.model_dump_json()
        loaded = AccessCatalog.model_validate_json(json_str)
        assert loaded.scope == ExtractionScope.STRUCTURE
        assert loaded.tables == []

    def test_json_roundtrip_with_table(self) -> None:
        catalog = make_minimal_catalog()
        catalog.tables.append(
            TableDef(
                name="TestTable",
                fields=[
                    FieldDef(
                        name="ID",
                        ordinal=0,
                        field_type=FieldType.LONG,
                        dao_type_code=4,
                        auto_number=True,
                        is_primary_key=True,
                    )
                ],
            )
        )
        json_str = catalog.model_dump_json()
        loaded = AccessCatalog.model_validate_json(json_str)
        assert len(loaded.tables) == 1
        assert loaded.tables[0].name == "TestTable"
        assert loaded.tables[0].fields[0].field_type == FieldType.LONG

    def test_dict_roundtrip(self) -> None:
        catalog = make_minimal_catalog()
        data = catalog.model_dump(mode="json")
        loaded = AccessCatalog.model_validate(data)
        assert loaded.extracted_by == "test"

    def test_scope_values(self) -> None:
        for scope in ExtractionScope:
            catalog = AccessCatalog(
                extracted_at=datetime(2026, 1, 1),
                extracted_by="test",
                scope=scope,
                properties=DatabaseProperties(file_path="C:\\test.accdb"),
            )
            assert catalog.scope == scope


class TestCatalogConvenienceMethods:
    def test_table_names(self, sample_catalog: AccessCatalog) -> None:
        names = sample_catalog.table_names()
        assert "Customers" in names
        assert "Orders" in names

    def test_query_names(self, sample_catalog: AccessCatalog) -> None:
        assert "qryCustomerOrders" in sample_catalog.query_names()

    def test_has_vba(self, sample_catalog: AccessCatalog) -> None:
        assert sample_catalog.has_vba()

    def test_total_vba_lines(self, sample_catalog: AccessCatalog) -> None:
        assert sample_catalog.total_vba_lines() == 7

    def test_error_count_zero(self, sample_catalog: AccessCatalog) -> None:
        assert sample_catalog.error_count() == 0

    def test_error_count_with_errors(self, sample_catalog: AccessCatalog) -> None:
        sample_catalog.warnings.append(
            ExtractionWarning(
                object_type="table",
                object_name="TestTable",
                warning_code="COM_ERROR",
                message="test error",
                severity=WarningSeverity.ERROR,
            )
        )
        assert sample_catalog.error_count() == 1


class TestFieldDef:
    def test_field_types_all_serializable(self) -> None:
        """All FieldType enum values should serialize/deserialize cleanly."""
        for ft in FieldType:
            fd = FieldDef(name="test", ordinal=0, field_type=ft, dao_type_code=0)
            data = fd.model_dump(mode="json")
            loaded = FieldDef.model_validate(data)
            assert loaded.field_type == ft

    def test_optional_fields_default_none(self) -> None:
        fd = FieldDef(name="x", ordinal=0, field_type=FieldType.TEXT, dao_type_code=10)
        assert fd.size is None
        assert fd.default_value is None
        assert fd.validation_rule is None


class TestQueryDef:
    def test_sql_preserved_verbatim(self) -> None:
        """SQL text must never be modified — stored exactly as extracted."""
        sql = "SELECT IIF([Active], 'Yes', 'No') FROM [My Table] WHERE ID > #1/1/2020#;"
        qd = QueryDef(name="q1", query_type=QueryType.SELECT, sql_text=sql)
        assert qd.sql_text == sql

    def test_circular_dependency_default_false(self) -> None:
        qd = QueryDef(name="q1", query_type=QueryType.SELECT, sql_text="SELECT 1")
        assert qd.in_circular_dependency is False


class TestVBAModule:
    def test_locked_module_has_none_source(self) -> None:
        m = VBAModule(
            name="TestModule",
            module_type=VBAModuleType.STANDARD,
            source_code=None,
            is_vba_locked=True,
        )
        assert m.source_code is None
        assert m.is_vba_locked is True
