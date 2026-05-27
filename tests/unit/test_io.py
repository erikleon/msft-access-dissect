"""Unit tests for catalog serialization/deserialization."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from access_dissect.catalog.models import AccessCatalog
from access_dissect.utils.io import catalog_to_json_str, load_catalog, save_catalog


class TestCatalogIO:
    def test_save_and_load_json(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        path = tmp_path / "catalog.json"
        save_catalog(sample_catalog, path, fmt="json")
        assert path.exists()
        loaded = load_catalog(path)
        assert loaded.properties.title == sample_catalog.properties.title
        assert len(loaded.tables) == len(sample_catalog.tables)

    def test_save_and_load_yaml(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        path = tmp_path / "catalog.yaml"
        save_catalog(sample_catalog, path, fmt="yaml")
        assert path.exists()
        loaded = load_catalog(path)
        assert loaded.properties.title == sample_catalog.properties.title

    def test_json_is_valid_json(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        path = tmp_path / "catalog.json"
        save_catalog(sample_catalog, path)
        content = path.read_text(encoding="utf-8")
        data = json.loads(content)
        assert "schema_version" in data
        assert "tables" in data

    def test_load_nonexistent_file_raises(self, tmp_path: Path) -> None:
        with pytest.raises(FileNotFoundError):
            load_catalog(tmp_path / "nonexistent.json")

    def test_catalog_to_json_str(self, sample_catalog: AccessCatalog) -> None:
        json_str = catalog_to_json_str(sample_catalog)
        data = json.loads(json_str)
        assert data["scope"] == "structure"

    def test_tables_survive_roundtrip(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        path = tmp_path / "catalog.json"
        save_catalog(sample_catalog, path)
        loaded = load_catalog(path)
        original_names = {t.name for t in sample_catalog.tables}
        loaded_names = {t.name for t in loaded.tables}
        assert original_names == loaded_names

    def test_field_types_survive_roundtrip(self, sample_catalog: AccessCatalog, tmp_path: Path) -> None:
        path = tmp_path / "catalog.json"
        save_catalog(sample_catalog, path)
        loaded = load_catalog(path)

        orig_fields = {
            (t.name, f.name): f.field_type
            for t in sample_catalog.tables
            for f in t.fields
        }
        loaded_fields = {
            (t.name, f.name): f.field_type
            for t in loaded.tables
            for f in t.fields
        }
        assert orig_fields == loaded_fields
