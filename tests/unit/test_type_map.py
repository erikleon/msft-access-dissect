"""Unit tests for DAO type mapping and field type constants."""

from __future__ import annotations

import pytest

from access_dissect.catalog.constants import (
    DAO_TYPE_MAP,
    dao_query_type_to_enum,
    dao_type_to_field_type,
    parse_relation_attributes,
)
from access_dissect.catalog.models import FieldType, QueryType


class TestDAOTypeMap:
    def test_all_14_field_types_covered(self) -> None:
        """All expected DAO type codes should be in the map."""
        expected_codes = {1, 2, 3, 4, 5, 6, 7, 8, 10, 11, 12, 15, 16, 17, 20}
        for code in expected_codes:
            result = dao_type_to_field_type(code)
            assert result != FieldType.UNKNOWN, f"DAO type {code} should have a known mapping"

    def test_unknown_code_returns_unknown(self) -> None:
        assert dao_type_to_field_type(999) == FieldType.UNKNOWN

    def test_boolean_is_code_1(self) -> None:
        assert dao_type_to_field_type(1) == FieldType.BOOLEAN

    def test_text_is_code_10(self) -> None:
        assert dao_type_to_field_type(10) == FieldType.TEXT

    def test_memo_is_code_12(self) -> None:
        assert dao_type_to_field_type(12) == FieldType.MEMO

    def test_ole_object_is_code_11(self) -> None:
        assert dao_type_to_field_type(11) == FieldType.OLE_OBJECT

    def test_attachment_is_code_17(self) -> None:
        assert dao_type_to_field_type(17) == FieldType.ATTACHMENT

    def test_guid_is_code_15(self) -> None:
        assert dao_type_to_field_type(15) == FieldType.GUID

    def test_date_is_code_8(self) -> None:
        assert dao_type_to_field_type(8) == FieldType.DATE


class TestDAOQueryTypeMap:
    def test_select_is_0(self) -> None:
        assert dao_query_type_to_enum(0) == QueryType.SELECT

    def test_pass_through_is_7(self) -> None:
        assert dao_query_type_to_enum(7) == QueryType.PASS_THROUGH

    def test_unknown_code(self) -> None:
        assert dao_query_type_to_enum(999) == QueryType.UNKNOWN

    def test_all_known_codes(self) -> None:
        for code in range(10):
            result = dao_query_type_to_enum(code)
            assert isinstance(result, QueryType)


class TestRelationAttributes:
    def test_no_flags(self) -> None:
        # Attributes=0 means one-to-many, NO referential integrity enforced
        # DB_DONTENFORCE = 2 is what disables RI; if not set, RI IS enforced
        flags = parse_relation_attributes(0)
        # one-to-one flag not set
        assert flags["one_to_one"] is False
        # cascade flags not set
        assert flags["cascade_update"] is False
        assert flags["cascade_delete"] is False

    def test_dont_enforce_flag(self) -> None:
        flags = parse_relation_attributes(2)  # DB_DONTENFORCE
        assert flags["enforce_integrity"] is False

    def test_cascade_update(self) -> None:
        flags = parse_relation_attributes(256)  # DB_CASCADE_UPDATE
        assert flags["cascade_update"] is True

    def test_cascade_delete(self) -> None:
        flags = parse_relation_attributes(4096)  # DB_CASCADE_DELETE
        assert flags["cascade_delete"] is True

    def test_one_to_one(self) -> None:
        flags = parse_relation_attributes(1)  # DB_UNIQUE
        assert flags["one_to_one"] is True

    def test_combined_flags(self) -> None:
        # RI enforced + cascade update + cascade delete
        flags = parse_relation_attributes(256 + 4096)
        assert flags["cascade_update"] is True
        assert flags["cascade_delete"] is True
