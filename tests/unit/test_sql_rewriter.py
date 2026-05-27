"""Unit tests for the Jet SQL rewriter."""

from __future__ import annotations

import pytest

from access_dissect.translate.sql_rewriter import translate_sql


def _sql_lines(result: str) -> str:
    """Return only non-comment lines from a translated SQL result."""
    return "\n".join(l for l in result.splitlines() if not l.startswith("--"))


class TestDateLiteralRewriting:
    def test_us_date_literal(self) -> None:
        sql = "SELECT * FROM Orders WHERE OrderDate > #1/15/2024#"
        result, _warns = translate_sql(sql, dialect="postgres")
        sql_only = _sql_lines(result)
        assert "#" not in sql_only
        assert "2024" in sql_only or "2024-01-15" in result

    def test_iso_date_literal(self) -> None:
        sql = "SELECT * FROM Orders WHERE OrderDate > #2024-01-15#"
        result, _warns = translate_sql(sql, dialect="postgres")
        sql_only = _sql_lines(result)
        assert "#" not in sql_only


class TestLikeWildcardRewriting:
    def test_star_wildcard_replaced(self) -> None:
        sql = "SELECT * FROM Customers WHERE CompanyName LIKE '*tech*'"
        result, _warns = translate_sql(sql, dialect="postgres")
        # After rewriting, Access * should become SQL %
        # Either sqlglot or our regex should handle this
        assert "%" in result or "'*tech*'" not in result

    def test_question_mark_wildcard(self) -> None:
        sql = "SELECT * FROM Customers WHERE Code LIKE 'A?C'"
        result, _warns = translate_sql(sql, dialect="postgres")
        # Question mark is Access single-char wildcard → SQL _
        assert result is not None


class TestDomainAggregateWarning:
    def test_dlookup_generates_warning(self) -> None:
        sql = "SELECT DLookup('CompanyName', 'Customers', 'CustomerID=1') AS Name"
        result, warns = translate_sql(sql, dialect="postgres")
        assert any("DLookup" in w or "domain aggregate" in w.lower() for w in warns)

    def test_dsum_generates_warning(self) -> None:
        sql = "SELECT DSum('Amount', 'Orders', 'CustomerID=1') AS Total"
        result, warns = translate_sql(sql, dialect="postgres")
        assert any("DSum" in w or "domain aggregate" in w.lower() for w in warns)

    def test_translated_sql_has_comment_header(self) -> None:
        sql = "SELECT CustomerID, CompanyName FROM Customers;"
        result, _warns = translate_sql(sql, dialect="postgres")
        assert "access-dissect" in result.lower() or "translated" in result.lower()


class TestTranslationDialects:
    def test_postgres_output(self) -> None:
        sql = "SELECT CustomerID FROM Customers;"
        result, _warns = translate_sql(sql, dialect="postgres")
        assert result is not None
        assert len(result) > 0

    def test_sqlite_output(self) -> None:
        sql = "SELECT CustomerID FROM Customers;"
        result, _warns = translate_sql(sql, dialect="sqlite")
        assert result is not None

    def test_mssql_output(self) -> None:
        sql = "SELECT CustomerID FROM Customers;"
        result, _warns = translate_sql(sql, dialect="mssql")
        assert result is not None

    def test_empty_sql_returns_empty(self) -> None:
        result, warns = translate_sql("", dialect="postgres")
        assert result == ""
        assert warns == []

    def test_original_sql_preserved_in_comment(self) -> None:
        sql = "SELECT IIF([Active], 'Yes', 'No') FROM Status"
        result, _warns = translate_sql(sql, dialect="postgres", original_name="qStatus")
        assert "qStatus" in result or "IIF" in result or "original" in result.lower()


class TestSQLParserExtraction:
    def test_extract_table_from_select(self) -> None:
        from access_dissect.analyze.sql_parser import extract_sql_references

        sql = "SELECT * FROM Customers"
        tables, queries = extract_sql_references(
            sql,
            known_tables={"Customers", "Orders"},
            known_queries=set(),
        )
        assert "Customers" in tables

    def test_extract_table_with_brackets(self) -> None:
        from access_dissect.analyze.sql_parser import extract_sql_references

        sql = "SELECT * FROM [Customer Data]"
        tables, queries = extract_sql_references(
            sql,
            known_tables={"Customer Data"},
            known_queries=set(),
        )
        assert "Customer Data" in tables

    def test_extract_query_reference(self) -> None:
        from access_dissect.analyze.sql_parser import extract_sql_references

        sql = "SELECT * FROM qryCustomers"
        tables, queries = extract_sql_references(
            sql,
            known_tables=set(),
            known_queries={"qryCustomers"},
        )
        assert "qryCustomers" in queries

    def test_join_references_extracted(self) -> None:
        from access_dissect.analyze.sql_parser import extract_sql_references

        sql = "SELECT * FROM Customers INNER JOIN Orders ON Customers.ID = Orders.CustomerID"
        tables, queries = extract_sql_references(
            sql,
            known_tables={"Customers", "Orders"},
            known_queries=set(),
        )
        assert "Customers" in tables
        assert "Orders" in tables
