"""
PropertyExtractor: extract database-level properties and VBA reference library list.
"""

from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path
from typing import Any

from access_dissect.catalog.models import (
    DatabaseProperties,
    ExtractionWarning,
    VBAReference,
    WarningSeverity,
)


def _safe_str(obj: Any, attr: str, default: str | None = None) -> str | None:
    try:
        val = getattr(obj, attr)
        return str(val) if val is not None else default
    except Exception:
        return default


def _safe_prop(props: Any, name: str) -> str | None:
    """Read a named property from an Access Properties collection."""
    try:
        return str(props(name).Value)
    except Exception:
        return None


def _extract_vba_references(app: Any) -> tuple[list[VBAReference], bool]:
    """
    Extract VBA project references and detect if the project is locked.

    Returns (references, vba_project_locked).
    """
    refs: list[VBAReference] = []
    locked = False

    if app is None:
        return refs, locked

    try:
        vbe = app.VBE
        project = vbe.VBProjects.Item(1)

        # Try to enumerate components first — this is what throws on a locked project
        try:
            _ = project.VBComponents.Count
        except Exception as e:
            msg = str(e).lower()
            # HRESULT 0x800A03EC or "permission denied" / "password"
            if "password" in msg or "permission" in msg or "800a03ec" in msg.replace(" ", ""):
                locked = True
                return refs, locked

        # Extract references (library dependencies)
        try:
            for ref in project.References:
                try:
                    refs.append(
                        VBAReference(
                            name=str(ref.Name),
                            description=str(ref.Description) if ref.Description else "",
                            guid=str(ref.GUID) if ref.GUID else None,
                            major=int(ref.Major) if ref.Major else None,
                            minor=int(ref.Minor) if ref.Minor else None,
                            full_path=str(ref.FullPath) if ref.FullPath else None,
                            is_broken=bool(ref.IsBroken),
                        )
                    )
                except Exception:
                    pass
        except Exception:
            pass

    except Exception:
        pass

    return refs, locked


def extract_properties(
    app: Any,
    db: Any,
    file_path: Path,
) -> tuple[DatabaseProperties, list[ExtractionWarning]]:
    """
    Extract database-level properties from the Access.Application and DAO CurrentDb().

    Returns (DatabaseProperties, warnings).
    """
    warnings: list[ExtractionWarning] = []

    # Basic file stats
    file_size = None
    try:
        file_size = os.path.getsize(file_path)
    except Exception:
        pass

    is_accdb = file_path.suffix.lower() == ".accdb"

    # Access/Jet version
    access_version = None
    jet_version = None
    try:
        access_version = str(app.Version)
    except Exception:
        pass
    try:
        jet_version = str(db.Version)
    except Exception:
        pass

    # DAO Properties collection for standard metadata
    props = None
    try:
        props = db.Properties
    except Exception:
        pass

    def _prop(name: str) -> str | None:
        if props is None:
            return None
        return _safe_prop(props, name)

    title = _prop("Title")
    subject = _prop("Subject")
    author = _prop("Author")
    company = _prop("Company")
    description = _prop("Description")

    # Date parsing
    created_date = None
    last_modified = None
    try:
        date_created_str = _prop("Date Created")
        if date_created_str:
            created_date = _parse_access_date(date_created_str)
    except Exception:
        pass
    try:
        last_updated_str = _prop("Date Modified")
        if last_updated_str:
            last_modified = _parse_access_date(last_updated_str)
    except Exception:
        pass

    # Workgroup / MDW file
    workgroup_path = None
    try:
        workgroup_path = _prop("Jet OLEDB:System database")
    except Exception:
        pass

    # VBA references and lock status
    vba_refs, vba_locked = _extract_vba_references(app)

    if vba_locked:
        warnings.append(
            ExtractionWarning(
                object_type="database",
                object_name="VBAProject",
                warning_code="VBA_LOCKED",
                message=(
                    "The VBA project is password-protected. "
                    "VBA source code cannot be extracted without the VBA project password. "
                    "This is a migration blocker — you must obtain the unlocked version."
                ),
                severity=WarningSeverity.ERROR,
            )
        )

    broken_refs = [r for r in vba_refs if r.is_broken]
    if broken_refs:
        warnings.append(
            ExtractionWarning(
                object_type="database",
                object_name="VBAProject",
                warning_code="BROKEN_VBA_REFERENCES",
                message=(
                    f"{len(broken_refs)} broken VBA reference(s): "
                    f"{[r.name for r in broken_refs]}. "
                    "These are missing COM libraries that VBA code depends on."
                ),
                severity=WarningSeverity.WARNING,
                details={"broken_references": [r.name for r in broken_refs]},
            )
        )

    db_props = DatabaseProperties(
        title=title,
        subject=subject,
        author=author,
        company=company,
        description=description,
        created_date=created_date,
        last_modified=last_modified,
        access_version=access_version,
        jet_version=jet_version,
        file_path=str(file_path.resolve()),
        file_size_bytes=file_size,
        is_accdb=is_accdb,
        vba_references=vba_refs,
        vba_project_locked=vba_locked,
        workgroup_db_path=workgroup_path,
    )
    return db_props, warnings


def _parse_access_date(date_str: str) -> datetime | None:
    """Try to parse various date formats Access may return."""
    formats = [
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%m/%d/%Y",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except ValueError:
            continue
    return None
