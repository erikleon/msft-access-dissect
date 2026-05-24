"""
AccessSQLParser: extract table and query references from Jet/ACE SQL.

Strategy:
  1. Primary: sqlglot.parse(sql, read="access") → walk AST for Table nodes
  2. Fallback: regex on FROM/JOIN/INTO/UPDATE/DELETE keywords (for malformed SQL)
  3. Resolve names against the known catalog name sets
"""

from __future__ import annotations

import re
from typing import Sequence


# Regex fallback: find identifiers after FROM, JOIN, INTO, UPDATE, DELETE FROM
# Use \w+ (no spaces) to avoid matching "Customers AS c" → "Customers AS"
_SQL_REF_PATTERN = re.compile(
    r"\b(?:FROM|JOIN|INTO|UPDATE)\s+"
    r"(?:\[([^\]]+)\]|([a-zA-Z_]\w*))",
    re.IGNORECASE,
)

_BRACKET_IDENTIFIER = re.compile(r"\[([^\]]+)\]")

# Identifiers that look like table refs but are SQL keywords
_SQL_KEYWORDS = frozenset({
    "select", "from", "where", "join", "left", "right", "inner", "outer",
    "on", "and", "or", "not", "in", "is", "null", "as", "by", "group",
    "order", "having", "distinct", "top", "all", "union", "intersect",
    "except", "set", "into", "values", "insert", "update", "delete",
    "create", "drop", "alter", "table", "index", "view", "procedure",
    "with", "case", "when", "then", "else", "end", "between", "like",
    "exists", "any", "some",
})


def _normalize_name(name: str) -> str:
    """Strip brackets and extra whitespace from an identifier."""
    name = name.strip()
    if name.startswith("[") and name.endswith("]"):
        name = name[1:-1]
    return name.strip()


def _extract_with_sqlglot(sql: str) -> set[str]:
    """Try to extract table references using sqlglot AST."""
    try:
        import sqlglot
        import sqlglot.expressions as exp

        tables: set[str] = set()
        try:
            parsed = sqlglot.parse(sql, read="access")
            for statement in parsed:
                if statement is None:
                    continue
                for table in statement.find_all(exp.Table):
                    name = table.name
                    if name and name.lower() not in _SQL_KEYWORDS:
                        tables.add(_normalize_name(name))
        except Exception:
            pass
        return tables
    except ImportError:
        return set()


def _extract_with_regex(sql: str) -> set[str]:
    """Fallback regex extraction of table references."""
    tables: set[str] = set()

    # First pass: extract all [bracketed identifiers] — Access uses these a lot
    for match in _BRACKET_IDENTIFIER.finditer(sql):
        name = match.group(1).strip()
        if name and name.lower() not in _SQL_KEYWORDS:
            tables.add(name)

    # Second pass: regex on keywords
    for match in _SQL_REF_PATTERN.finditer(sql):
        # Group 1: bracketed, group 2: unbracketed
        name = (match.group(1) or match.group(2) or "").strip()
        if name and name.lower() not in _SQL_KEYWORDS:
            tables.add(_normalize_name(name))

    return tables


def extract_sql_references(
    sql: str,
    known_tables: set[str] | None = None,
    known_queries: set[str] | None = None,
) -> tuple[list[str], list[str]]:
    """
    Extract table and query references from an Access SQL statement.

    Returns (table_refs, query_refs) — lists of names found in the catalog.
    If known_tables/known_queries are None, returns all found names in table_refs.
    """
    if not sql or not sql.strip():
        return [], []

    # Try sqlglot first, fall back to regex
    refs = _extract_with_sqlglot(sql)
    if not refs:
        refs = _extract_with_regex(sql)

    if known_tables is None and known_queries is None:
        return sorted(refs), []

    table_refs: list[str] = []
    query_refs: list[str] = []

    for ref in refs:
        # Case-insensitive lookup against known names
        ref_lower = ref.lower()
        matched = False
        if known_tables:
            for t in known_tables:
                if t.lower() == ref_lower:
                    table_refs.append(t)
                    matched = True
                    break
        if not matched and known_queries:
            for q in known_queries:
                if q.lower() == ref_lower:
                    query_refs.append(q)
                    matched = True
                    break

    return table_refs, query_refs
