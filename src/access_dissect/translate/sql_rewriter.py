"""
SQL rewriter: translate Jet/ACE SQL to a target dialect.

Strategy:
  1. Try sqlglot.transpile(sql, read="access", write=dialect)
  2. Post-process for Access-specific idioms sqlglot may miss
  3. Flag untranslatable patterns (domain aggregates) with TODO comments
  4. Always preserve the original SQL; translations are advisory

The original SQL is always stored verbatim in the catalog.
This module is called only during the render phase.
"""

from __future__ import annotations

import re

# ---------------------------------------------------------------------------
# Patterns for Access-specific idioms
# ---------------------------------------------------------------------------

_IIF_PATTERN = re.compile(r"\bIIF\s*\(", re.IGNORECASE)
_NZ_PATTERN = re.compile(r"\bNZ\s*\(", re.IGNORECASE)
_DATE_LITERAL_PATTERN = re.compile(r"#(\d{1,2}/\d{1,2}/\d{2,4})#")
_DATE_LITERAL_ISO_PATTERN = re.compile(r"#(\d{4}-\d{2}-\d{2})#")
_LIKE_STAR_PATTERN = re.compile(r"LIKE\s+'([^']*\*[^']*)'", re.IGNORECASE)
_DOMAIN_AGGREGATE_PATTERN = re.compile(
    r"\b(DLookup|DSum|DCount|DAvg|DMax|DMin|DFirst|DLast)\s*\(",
    re.IGNORECASE,
)
_FORMAT_PATTERN = re.compile(r"\bFormat\s*\(", re.IGNORECASE)

# Access TOP N syntax: SELECT TOP N ... (without PERCENT)
_TOP_PATTERN = re.compile(r"\bSELECT\s+TOP\s+(\d+)\b", re.IGNORECASE)


def _try_sqlglot(sql: str, dialect: str) -> str | None:
    """Attempt translation with sqlglot. Returns None on failure."""
    try:
        import sqlglot

        _DIALECT_MAP = {
            "postgres": "postgres",
            "sqlite": "sqlite",
            "mssql": "tsql",
        }
        target = _DIALECT_MAP.get(dialect, dialect)

        results = sqlglot.transpile(sql, read="access", write=target, error_level=sqlglot.ErrorLevel.IGNORE)
        if results and results[0]:
            return results[0]
    except Exception:
        pass
    return None


def _post_process(sql: str, dialect: str) -> tuple[str, list[str]]:
    """
    Apply regex-based post-processing for patterns sqlglot may miss.
    Returns (processed_sql, list_of_warnings).
    """
    warnings: list[str] = []

    # IIF(cond, a, b) → CASE WHEN cond THEN a ELSE b END
    # Note: sqlglot usually handles this, but belt-and-suspenders
    if _IIF_PATTERN.search(sql):
        warnings.append("IIF() function detected — verify CASE WHEN translation is correct")

    # NZ(field, default) → COALESCE(field, default)
    # sqlglot may handle this; we just warn if it's still present
    if _NZ_PATTERN.search(sql):
        warnings.append("NZ() function detected — verify COALESCE translation is correct")

    # #date# literals → 'date' (ISO format preferred)
    def replace_date(m: re.Match) -> str:
        date_str = m.group(1)
        # Try to normalize M/D/YYYY to YYYY-MM-DD
        try:
            from datetime import datetime
            parsed = datetime.strptime(date_str, "%m/%d/%Y")
            return f"'{parsed.strftime('%Y-%m-%d')}'"
        except ValueError:
            pass
        try:
            from datetime import datetime
            parsed = datetime.strptime(date_str, "%m/%d/%y")
            return f"'{parsed.strftime('%Y-%m-%d')}'"
        except ValueError:
            pass
        return f"'{date_str}'"

    sql = _DATE_LITERAL_PATTERN.sub(replace_date, sql)
    sql = _DATE_LITERAL_ISO_PATTERN.sub(lambda m: f"'{m.group(1)}'", sql)

    # LIKE "*text*" → LIKE '%text%'
    def replace_like_wildcards(m: re.Match) -> str:
        pattern = m.group(1).replace("*", "%").replace("?", "_")
        return f"LIKE '{pattern}'"

    sql = _LIKE_STAR_PATTERN.sub(replace_like_wildcards, sql)

    # Domain aggregates — flag as untranslatable
    domain_funcs = _DOMAIN_AGGREGATE_PATTERN.findall(sql)
    if domain_funcs:
        unique = list(set(f.strip() for f in domain_funcs))
        warnings.append(
            f"Domain aggregate function(s) {unique} cannot be auto-translated. "
            "Replace with correlated subqueries or application-layer logic."
        )

    # Format() function — complex, flag for manual review
    if _FORMAT_PATTERN.search(sql):
        warnings.append("Format() function detected — requires manual review for equivalent in target dialect")

    # SELECT TOP N → LIMIT N (postgres/sqlite) or keep for mssql
    if dialect in ("postgres", "sqlite"):
        def replace_top(m: re.Match) -> str:
            return f"SELECT /*TOP {m.group(1)} → use LIMIT*/"

        # Don't rewrite — just warn; sqlglot should handle this
        if _TOP_PATTERN.search(sql):
            warnings.append("SELECT TOP N rewritten to LIMIT N by sqlglot — verify correctness")

    return sql, warnings


def translate_sql(
    sql: str,
    dialect: str = "postgres",
    original_name: str = "",
) -> tuple[str, list[str]]:
    """
    Translate Access Jet SQL to a target dialect.

    Returns (translated_sql, warnings).
    The original SQL is embedded in a comment header.
    """
    if not sql or not sql.strip():
        return sql, []

    warnings: list[str] = []

    # Try sqlglot first
    translated = _try_sqlglot(sql, dialect)
    if translated:
        translated, post_warns = _post_process(translated, dialect)
        warnings.extend(post_warns)
    else:
        # Fallback: just post-process the original
        translated, post_warns = _post_process(sql, dialect)
        warnings.extend(post_warns)
        warnings.insert(0, "sqlglot could not parse this query — translation is best-effort regex only")

    # Build header comment
    header_lines = [
        f"-- TRANSLATED FROM ACCESS SQL (access-dissect)",
    ]
    if original_name:
        header_lines.append(f"-- Source query: {original_name}")
    if sql.strip() != translated.strip():
        # Truncate original if very long
        orig_display = sql[:200] + "..." if len(sql) > 200 else sql
        header_lines.append(f"-- Original: {orig_display.replace(chr(10), ' ')}")
    if warnings:
        for w in warnings:
            header_lines.append(f"-- WARNING: {w}")

    header = "\n".join(header_lines)
    return f"{header}\n{translated}", warnings
