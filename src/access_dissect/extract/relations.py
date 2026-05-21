"""
RelationExtractor: extract relationship definitions from CurrentDb().Relations.
"""

from __future__ import annotations

from typing import Any

from access_dissect.catalog.constants import parse_relation_attributes
from access_dissect.catalog.models import (
    ExtractionWarning,
    RelationDef,
    RelationField,
    WarningSeverity,
)


def _safe_str(obj: Any, attr: str, default: str = "") -> str:
    try:
        val = getattr(obj, attr)
        return str(val) if val else default
    except Exception:
        return default


def extract_relation(relation: Any) -> tuple[RelationDef | None, list[ExtractionWarning]]:
    """Extract a single RelationDef from a DAO Relation COM object."""
    warnings: list[ExtractionWarning] = []
    name = ""
    try:
        name = _safe_str(relation, "Name")
    except Exception:
        return None, warnings

    try:
        attrs = int(relation.Attributes)
    except Exception:
        attrs = 2  # default: no RI enforcement

    attr_flags = parse_relation_attributes(attrs)

    parent_table = _safe_str(relation, "Table")
    child_table = _safe_str(relation, "ForeignTable")

    # Extract field pairs
    rel_fields: list[RelationField] = []
    try:
        for f in relation.Fields:
            try:
                rel_fields.append(
                    RelationField(
                        child_field=str(f.ForeignName),
                        parent_field=str(f.Name),
                    )
                )
            except Exception:
                pass
    except Exception as e:
        warnings.append(
            ExtractionWarning(
                object_type="relation",
                object_name=name,
                warning_code="RELATION_FIELDS_ERROR",
                message=f"Could not enumerate relation fields: {e}",
                severity=WarningSeverity.WARNING,
            )
        )

    if not attr_flags["enforce_integrity"]:
        warnings.append(
            ExtractionWarning(
                object_type="relation",
                object_name=name,
                warning_code="NO_RI_ENFORCEMENT",
                message=(
                    f"Relation '{name}' ({parent_table} → {child_table}) does not enforce "
                    "referential integrity. Consider adding FK constraints in the target database."
                ),
                severity=WarningSeverity.INFO,
            )
        )

    rel = RelationDef(
        name=name,
        parent_table=parent_table,
        child_table=child_table,
        fields=rel_fields,
        enforce_integrity=attr_flags["enforce_integrity"],
        cascade_update=attr_flags["cascade_update"],
        cascade_delete=attr_flags["cascade_delete"],
        one_to_one=attr_flags["one_to_one"],
    )
    return rel, warnings


def extract_all_relations(db: Any) -> tuple[list[RelationDef], list[ExtractionWarning]]:
    """Extract all relations from CurrentDb().Relations."""
    relations: list[RelationDef] = []
    all_warnings: list[ExtractionWarning] = []

    try:
        rels = db.Relations
    except Exception as e:
        all_warnings.append(
            ExtractionWarning(
                object_type="database",
                object_name="Relations",
                warning_code="COM_ERROR",
                message=f"Could not access Relations collection: {e}",
                severity=WarningSeverity.ERROR,
            )
        )
        return relations, all_warnings

    count = 0
    try:
        count = rels.Count
    except Exception:
        pass

    for i in range(count):
        name = f"<index {i}>"
        try:
            rel = rels(i)
            try:
                name = _safe_str(rel, "Name")
            except Exception:
                pass
            rel_def, warnings = extract_relation(rel)
            all_warnings.extend(warnings)
            if rel_def is not None:
                # Skip internal MSys relations
                if not rel_def.name.startswith("MSys"):
                    relations.append(rel_def)
        except Exception as e:
            all_warnings.append(
                ExtractionWarning(
                    object_type="relation",
                    object_name=name,
                    warning_code="COM_ERROR",
                    message=f"Unexpected error extracting relation: {e}",
                    severity=WarningSeverity.ERROR,
                )
            )

    return relations, all_warnings
