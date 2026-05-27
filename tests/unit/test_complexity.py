"""Unit tests for the complexity scoring rules."""

from __future__ import annotations

import pytest

from access_dissect.analyze.complexity import (
    ObjectScore,
    score_form,
    score_query,
    score_table,
    score_vba_module,
)
from access_dissect.catalog.models import (
    ControlDef,
    FieldDef,
    FieldType,
    FormDef,
    IndexDef,
    IndexField,
    QueryDef,
    QueryType,
    TableDef,
    VBAModule,
    VBAModuleType,
    VBAReference,
)


class TestTableScoring:
    def test_clean_table_scores_zero(self) -> None:
        table = TableDef(
            name="Clean",
            fields=[
                FieldDef(name="ID", ordinal=0, field_type=FieldType.LONG, dao_type_code=4,
                         is_primary_key=True),
                FieldDef(name="Name", ordinal=1, field_type=FieldType.TEXT, dao_type_code=10,
                         size=50),
            ],
            indexes=[IndexDef(name="PrimaryKey", fields=[IndexField(field_name="ID")], primary=True)],
        )
        score = score_table(table)
        assert score.score == 0.0
        assert score.band == "low"

    def test_no_pk_adds_score(self) -> None:
        table = TableDef(
            name="NoPK",
            fields=[FieldDef(name="Name", ordinal=0, field_type=FieldType.TEXT, dao_type_code=10)],
        )
        score = score_table(table)
        assert score.score >= 1.0
        assert any("primary key" in r.lower() for r in score.reasons)

    def test_ole_field_adds_score(self) -> None:
        table = TableDef(
            name="WithOLE",
            fields=[
                FieldDef(name="ID", ordinal=0, field_type=FieldType.LONG, dao_type_code=4,
                         is_primary_key=True),
                FieldDef(name="Picture", ordinal=1, field_type=FieldType.OLE_OBJECT, dao_type_code=11),
            ],
        )
        score = score_table(table)
        assert score.score >= 1.0

    def test_fields_with_spaces_add_score(self) -> None:
        table = TableDef(
            name="SpaceTable",
            fields=[
                FieldDef(name="First Name", ordinal=0, field_type=FieldType.TEXT, dao_type_code=10),
                FieldDef(name="Last Name", ordinal=1, field_type=FieldType.TEXT, dao_type_code=10),
            ],
        )
        score = score_table(table)
        assert score.score > 0

    def test_score_capped_at_10(self) -> None:
        # Table with everything bad
        fields = [
            FieldDef(name=f"Field {i}", ordinal=i, field_type=FieldType.OLE_OBJECT, dao_type_code=11)
            for i in range(20)
        ]
        table = TableDef(name="Nightmare", fields=fields)
        score = score_table(table)
        assert score.score <= 10.0


class TestQueryScoring:
    def test_pass_through_scores_high(self) -> None:
        query = QueryDef(
            name="qPassThru",
            query_type=QueryType.PASS_THROUGH,
            sql_text="SELECT * FROM dbo.Customers",
            is_pass_through=True,
        )
        score = score_query(query)
        assert score.score >= 2.0
        assert any("pass-through" in r.lower() for r in score.reasons)

    def test_dlookup_scores_high(self) -> None:
        query = QueryDef(
            name="qDomain",
            query_type=QueryType.SELECT,
            sql_text="SELECT DLookup('Name', 'Customers', 'ID=1') AS N FROM Dummy",
        )
        score = score_query(query)
        assert score.score >= 1.5

    def test_circular_adds_score(self) -> None:
        query = QueryDef(
            name="qCircular",
            query_type=QueryType.SELECT,
            sql_text="SELECT * FROM qOther",
            in_circular_dependency=True,
        )
        score = score_query(query)
        assert score.score >= 1.0

    def test_simple_select_scores_zero(self) -> None:
        query = QueryDef(
            name="qSimple",
            query_type=QueryType.SELECT,
            sql_text="SELECT CustomerID, CompanyName FROM Customers",
        )
        score = score_query(query)
        assert score.score == 0.0


class TestFormScoring:
    def test_form_with_vba_over_200_lines(self) -> None:
        form = FormDef(
            name="frmComplex",
            has_vba_module=True,
            vba_module_name="Form_frmComplex",
        )
        modules = [
            VBAModule(
                name="Form_frmComplex",
                module_type=VBAModuleType.FORM_MODULE,
                source_code="\n".join(["' code"] * 250),
                line_count=250,
            )
        ]
        score = score_form(form, modules)
        assert score.score >= 3.0

    def test_form_with_activex_scores_extra(self) -> None:
        form = FormDef(
            name="frmActiveX",
            controls=[
                ControlDef(
                    name="ctrl1",
                    control_type=119,
                    control_type_name="ActiveX",
                    is_active_x=True,
                    active_x_prog_id="SomeControl.1",
                )
            ],
        )
        score = score_form(form, [])
        assert score.score >= 1.0

    def test_clean_form_scores_zero(self) -> None:
        form = FormDef(name="frmSimple")
        score = score_form(form, [])
        assert score.score == 0.0


class TestObjectScoreBands:
    def test_band_low(self) -> None:
        s = ObjectScore(object_type="table", object_name="T", score=2.5)
        assert s.band == "low"

    def test_band_medium(self) -> None:
        s = ObjectScore(object_type="table", object_name="T", score=5.0)
        assert s.band == "medium"

    def test_band_high(self) -> None:
        s = ObjectScore(object_type="table", object_name="T", score=7.5)
        assert s.band == "high"
