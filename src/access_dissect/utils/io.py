"""
Catalog serialization / deserialization utilities.

Supports JSON (default) and YAML output.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import yaml

from access_dissect.catalog.models import AccessCatalog


def _json_default(obj: Any) -> Any:
    """Custom JSON serializer for types Pydantic doesn't handle natively."""
    from datetime import datetime

    if isinstance(obj, datetime):
        return obj.isoformat()
    raise TypeError(f"Object of type {type(obj).__name__} is not JSON serializable")


def save_catalog(catalog: AccessCatalog, path: Path, fmt: str = "json") -> None:
    """Write an AccessCatalog to disk as JSON or YAML."""
    path = Path(path)
    data = catalog.model_dump(mode="json")

    if fmt == "yaml":
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8") as fh:
            yaml.dump(data, fh, default_flow_style=False, allow_unicode=True, sort_keys=False)
    else:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2, ensure_ascii=False, default=_json_default)


def load_catalog(path: Path) -> AccessCatalog:
    """Load an AccessCatalog from a JSON or YAML file."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Catalog file not found: {path}")

    with path.open("r", encoding="utf-8") as fh:
        if path.suffix.lower() in (".yaml", ".yml"):
            data = yaml.safe_load(fh)
        else:
            data = json.load(fh)

    return AccessCatalog.model_validate(data)


def catalog_to_json_str(catalog: AccessCatalog, indent: int = 2) -> str:
    """Serialize a catalog to a JSON string."""
    return catalog.model_dump_json(indent=indent)
