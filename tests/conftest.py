"""
Shared pytest fixtures for access-dissect tests.

Provides:
  - Mock COM objects (no Access required)
  - Sample catalog fixture
  - pytest marks
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any
from unittest.mock import MagicMock

import pytest

from access_dissect.catalog.models import (
    AccessCatalog,
    DatabaseProperties,
    ExtractionScope,
    FieldDef,
    FieldType,
    FormDef,
    IndexDef,
    IndexField,
    QueryDef,
    QueryType,
    RelationDef,
    RelationField,
    TableDef,
    VBAModule,
    VBAModuleType,
    VBAReference,
)


# ---------------------------------------------------------------------------
# Mark: requires_access
# ---------------------------------------------------------------------------


def pytest_configure(config: Any) -> None:
    config.addinivalue_line(
        "markers",
        "requires_access: marks tests that need Microsoft Access installed on Windows",
    )


# ---------------------------------------------------------------------------
# Mock COM builder helpers
# ---------------------------------------------------------------------------


def _mock_field(
    name: str,
    ordinal: int,
    field_type: int = 10,
    size: int = 255,
    required: bool = False,
    default: str | None = None,
    auto_number: bool = False,
) -> MagicMock:
    f = MagicMock()
    f.Name = name
    f.Type = field_type
    f.Size = size
    f.Required = required
    f.AllowZeroLength = True
    f.DefaultValue = default
    f.ValidationRule = None
    f.ValidationText = None
    f.Description = None
    f.AutoNumber = auto_number
    f.Attributes = 0
    f.Properties = []
    return f


def _mock_index(name: str, fields: list[str], primary: bool = False, unique: bool = False) -> MagicMock:
    idx = MagicMock()
    idx.Name = name
    idx.Primary = primary
    idx.Unique = unique
    idx.Required = primary
    idx.IgnoreNulls = False
    idx.Foreign = False

    mock_fields = []
    for fname in fields:
        mf = MagicMock()
        mf.Name = fname
        mf.Attributes = 0
        mock_fields.append(mf)
    idx.Fields = mock_fields
    return idx


def _iterable_mock(items: list) -> MagicMock:
    """Make a MagicMock that is iterable and has .Count and __call__(i) interface."""
    m = MagicMock()
    m.Count = len(items)
    m.__iter__ = lambda self: iter(items)
    m.__call__ = lambda i: items[i]
    m.__getitem__ = lambda self, i: items[i]
    return m


@pytest.fixture
def mock_tabledef_simple() -> MagicMock:
    """A simple Access TableDef mock with 3 fields and one PK index."""
    td = MagicMock()
    td.Name = "Customers"
    td.Attributes = 0  # regular table, not system, not hidden, not linked
    td.Connect = ""
    td.SourceTableName = ""
    td.Description = "Customer table"

    fields = [
        _mock_field("CustomerID", 0, field_type=4, auto_number=True),
        _mock_field("CompanyName", 1, field_type=10, size=50, required=True),
        _mock_field("ContactName", 2, field_type=10, size=30),
    ]
    td.Fields = _iterable_mock(fields)

    pk_index = _mock_index("PrimaryKey", ["CustomerID"], primary=True, unique=True)
    td.Indexes = _iterable_mock([pk_index])

    return td


@pytest.fixture
def mock_querydef_simple() -> MagicMock:
    """A simple SELECT query mock."""
    qd = MagicMock()
    qd.Name = "qryCustomers"
    qd.Type = 0  # dbQSelect
    qd.SQL = "SELECT CustomerID, CompanyName FROM Customers ORDER BY CompanyName;"
    qd.Parameters = _iterable_mock([])
    qd.ReturnsRecords = True
    qd.Connect = ""
    return qd


@pytest.fixture
def mock_database(mock_tabledef_simple: MagicMock, mock_querydef_simple: MagicMock) -> MagicMock:
    """A mock DAO Database object with tables and queries."""
    db = MagicMock()
    db.TableDefs = _iterable_mock([mock_tabledef_simple])
    db.QueryDefs = _iterable_mock([mock_querydef_simple])
    db.Relations = _iterable_mock([])
    db.Version = "16.0"

    props_mock = MagicMock()
    props_mock.side_effect = lambda name: MagicMock(Value=None)
    db.Properties = MagicMock()
    db.Properties.__call__ = props_mock
    return db


# ---------------------------------------------------------------------------
# Sample catalog fixture
# ---------------------------------------------------------------------------


@pytest.fixture
def sample_catalog() -> AccessCatalog:
    """A minimal but complete AccessCatalog for testing renderers and analyzers."""
    return AccessCatalog(
        extracted_at=datetime(2026, 1, 15, 10, 30, 0),
        extracted_by="access-dissect 0.1.0",
        scope=ExtractionScope.STRUCTURE,
        properties=DatabaseProperties(
            title="Northwind",
            author="Microsoft",
            file_path="C:\\Users\\test\\Northwind.accdb",
            file_size_bytes=1_234_567,
            is_accdb=True,
            access_version="16.0",
            vba_references=[
                VBAReference(name="VBA", description="Visual Basic For Applications"),
                VBAReference(name="Access", description="Microsoft Access 16.0 Object Library"),
                VBAReference(name="DAO", description="Microsoft Office 16.0 Access database engine"),
            ],
        ),
        tables=[
            TableDef(
                name="Customers",
                fields=[
                    FieldDef(name="CustomerID", ordinal=0, field_type=FieldType.LONG,
                             dao_type_code=4, auto_number=True, is_primary_key=True),
                    FieldDef(name="CompanyName", ordinal=1, field_type=FieldType.TEXT,
                             dao_type_code=10, size=50, required=True),
                    FieldDef(name="Country", ordinal=2, field_type=FieldType.TEXT,
                             dao_type_code=10, size=15),
                ],
                indexes=[
                    IndexDef(
                        name="PrimaryKey",
                        fields=[IndexField(field_name="CustomerID")],
                        primary=True,
                        unique=True,
                    )
                ],
            ),
            TableDef(
                name="Orders",
                fields=[
                    FieldDef(name="OrderID", ordinal=0, field_type=FieldType.LONG,
                             dao_type_code=4, auto_number=True, is_primary_key=True),
                    FieldDef(name="CustomerID", ordinal=1, field_type=FieldType.LONG,
                             dao_type_code=4, required=True),
                    FieldDef(name="OrderDate", ordinal=2, field_type=FieldType.DATE,
                             dao_type_code=8),
                ],
                indexes=[
                    IndexDef(
                        name="PrimaryKey",
                        fields=[IndexField(field_name="OrderID")],
                        primary=True,
                        unique=True,
                    )
                ],
            ),
        ],
        queries=[
            QueryDef(
                name="qryCustomerOrders",
                query_type=QueryType.SELECT,
                sql_text=(
                    "SELECT c.CompanyName, o.OrderID, o.OrderDate "
                    "FROM Customers AS c "
                    "INNER JOIN Orders AS o ON c.CustomerID = o.CustomerID "
                    "ORDER BY c.CompanyName, o.OrderDate DESC;"
                ),
            ),
        ],
        relations=[
            RelationDef(
                name="OrdersCustomers",
                parent_table="Customers",
                child_table="Orders",
                fields=[RelationField(parent_field="CustomerID", child_field="CustomerID")],
                enforce_integrity=True,
                cascade_update=True,
                cascade_delete=False,
            )
        ],
        modules=[
            VBAModule(
                name="modUtilities",
                module_type=VBAModuleType.STANDARD,
                source_code=(
                    "Option Compare Database\nOption Explicit\n\n"
                    "Function FormatPhone(phone As String) As String\n"
                    "    FormatPhone = Format(phone, \"(###) ###-####\")\n"
                    "End Function\n"
                ),
                line_count=7,
                declaration_lines=2,
            )
        ],
    )


# ---------------------------------------------------------------------------
# Fixtures directory path
# ---------------------------------------------------------------------------


@pytest.fixture
def fixtures_dir() -> Path:
    return Path(__file__).parent / "fixtures"
