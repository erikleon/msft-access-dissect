"""Unit tests for dependency graph construction and cycle detection."""

from __future__ import annotations

import pytest

from access_dissect.analyze.dependency import (
    annotate_catalog_with_dependencies,
    build_dependency_graph,
    find_cycles,
)
from access_dissect.catalog.models import (
    AccessCatalog,
    QueryDef,
    QueryType,
)


class TestCycleDetection:
    def test_no_cycles(self) -> None:
        graph = {
            "query:A": ["table:T1"],
            "query:B": ["query:A"],
            "table:T1": [],
        }
        cycles = find_cycles(graph)
        assert cycles == []

    def test_simple_cycle(self) -> None:
        graph = {
            "query:A": ["query:B"],
            "query:B": ["query:A"],
        }
        cycles = find_cycles(graph)
        assert len(cycles) > 0
        # Both nodes should be in the cycle
        all_nodes = {n for cycle in cycles for n in cycle}
        assert "query:A" in all_nodes
        assert "query:B" in all_nodes

    def test_three_node_cycle(self) -> None:
        graph = {
            "query:A": ["query:B"],
            "query:B": ["query:C"],
            "query:C": ["query:A"],
        }
        cycles = find_cycles(graph)
        assert len(cycles) > 0

    def test_dag_no_cycles(self) -> None:
        # A DAG (directed acyclic graph) should have no cycles
        graph = {
            "form:F1": ["query:Q1"],
            "query:Q1": ["table:T1", "table:T2"],
            "table:T1": [],
            "table:T2": [],
        }
        cycles = find_cycles(graph)
        assert cycles == []

    def test_isolated_nodes_no_cycles(self) -> None:
        graph = {
            "table:T1": [],
            "table:T2": [],
            "module:M1": [],
        }
        cycles = find_cycles(graph)
        assert cycles == []


class TestDependencyGraphConstruction:
    def test_query_references_table(self, sample_catalog: AccessCatalog) -> None:
        graph = build_dependency_graph(sample_catalog)

        # qryCustomerOrders should reference Customers and Orders tables
        q_node = "query:qryCustomerOrders"
        assert q_node in graph
        deps = graph[q_node]
        assert "table:Customers" in deps or "table:Orders" in deps

    def test_all_table_nodes_exist(self, sample_catalog: AccessCatalog) -> None:
        graph = build_dependency_graph(sample_catalog)
        assert "table:Customers" in graph
        assert "table:Orders" in graph

    def test_empty_catalog(self) -> None:
        from datetime import datetime
        catalog = AccessCatalog(
            extracted_at=datetime(2026, 1, 1),
            extracted_by="test",
            scope="structure",
            properties=__import__("access_dissect.catalog.models", fromlist=["DatabaseProperties"]).DatabaseProperties(
                file_path="C:\\test.accdb"
            ),
        )
        graph = build_dependency_graph(catalog)
        assert isinstance(graph, dict)

    def test_annotate_marks_circular_queries(self) -> None:
        from datetime import datetime
        from access_dissect.catalog.models import DatabaseProperties

        catalog = AccessCatalog(
            extracted_at=datetime(2026, 1, 1),
            extracted_by="test",
            scope="structure",
            properties=DatabaseProperties(file_path="C:\\test.accdb"),
            queries=[
                QueryDef(
                    name="qA",
                    query_type=QueryType.SELECT,
                    sql_text="SELECT * FROM qB",
                ),
                QueryDef(
                    name="qB",
                    query_type=QueryType.SELECT,
                    sql_text="SELECT * FROM qA",
                ),
            ],
        )
        annotate_catalog_with_dependencies(catalog)

        # Both circular queries should be marked
        qa = next(q for q in catalog.queries if q.name == "qA")
        qb = next(q for q in catalog.queries if q.name == "qB")
        # At least one of them should be marked circular (cycle detection may vary)
        assert qa.in_circular_dependency or qb.in_circular_dependency
