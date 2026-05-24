"""
DependencyAnalyzer: build a directed dependency graph and detect cycles.

Graph node keys: "type:name"
  e.g. "table:Customers", "query:qryOrders", "form:frmCustomers", "module:modUtils"

Graph edges: source depends on target
  query → table/query
  form → table/query/form (via record_source and subforms)
  report → table/query/report
  module → form/query/table (via DoCmd and SQL string patterns in VBA)

Cycle detection uses Kahn's algorithm (BFS topological sort).
"""

from __future__ import annotations

import re
from collections import defaultdict, deque

from access_dissect.analyze.sql_parser import extract_sql_references
from access_dissect.catalog.models import AccessCatalog, QueryDef


# VBA patterns for DoCmd calls that reference other objects
_VBA_OPEN_PATTERNS = [
    # DoCmd.OpenForm "FormName"
    (re.compile(r'DoCmd\.OpenForm\s+"([^"]+)"', re.IGNORECASE), "form"),
    # DoCmd.OpenReport "ReportName"
    (re.compile(r'DoCmd\.OpenReport\s+"([^"]+)"', re.IGNORECASE), "report"),
    # DoCmd.OpenQuery "QueryName"
    (re.compile(r'DoCmd\.OpenQuery\s+"([^"]+)"', re.IGNORECASE), "query"),
    # DoCmd.OpenTable "TableName"
    (re.compile(r'DoCmd\.OpenTable\s+"([^"]+)"', re.IGNORECASE), "table"),
]


def _node(obj_type: str, name: str) -> str:
    return f"{obj_type}:{name}"


def build_dependency_graph(catalog: AccessCatalog) -> dict[str, list[str]]:
    """
    Build the full dependency graph for the catalog.

    Returns an adjacency list: {node: [list of nodes it depends on]}.
    Also updates query.references_tables and query.references_queries in-place.
    """
    graph: dict[str, list[str]] = defaultdict(list)

    known_tables = catalog.table_names()
    known_queries = catalog.query_names()
    known_forms = catalog.form_names()
    known_reports = catalog.report_names()

    # Initialize all nodes
    for t in catalog.tables:
        graph[_node("table", t.name)]
    for q in catalog.queries:
        graph[_node("query", q.name)]
    for f in catalog.forms:
        graph[_node("form", f.name)]
    for r in catalog.reports:
        graph[_node("report", r.name)]
    for m in catalog.modules:
        graph[_node("module", m.name)]

    # --- Queries → tables/queries via SQL ---
    for query in catalog.queries:
        src = _node("query", query.name)
        t_refs, q_refs = extract_sql_references(
            query.sql_text,
            known_tables=known_tables,
            known_queries=known_queries,
        )
        query.references_tables = t_refs
        query.references_queries = q_refs

        for t in t_refs:
            dep = _node("table", t)
            if dep not in graph[src]:
                graph[src].append(dep)
        for q in q_refs:
            dep = _node("query", q)
            if dep != src and dep not in graph[src]:
                graph[src].append(dep)

    # --- Forms → tables/queries via record_source; forms via subforms ---
    for form in catalog.forms:
        src = _node("form", form.name)
        if form.record_source:
            rs = form.record_source.strip()
            if rs.lower() in {q.lower() for q in known_queries}:
                dep = _node("query", rs)
                if dep not in graph[src]:
                    graph[src].append(dep)
            elif rs.lower() in {t.lower() for t in known_tables}:
                dep = _node("table", rs)
                if dep not in graph[src]:
                    graph[src].append(dep)
            elif rs.upper().startswith("SELECT"):
                # Inline SQL in record source
                t_refs, q_refs = extract_sql_references(
                    rs, known_tables=known_tables, known_queries=known_queries
                )
                for t in t_refs:
                    dep = _node("table", t)
                    if dep not in graph[src]:
                        graph[src].append(dep)
                for q in q_refs:
                    dep = _node("query", q)
                    if dep not in graph[src]:
                        graph[src].append(dep)

        # Subform dependencies
        for sf_src in form.subforms:
            if sf_src.lower() in {f.lower() for f in known_forms}:
                dep = _node("form", sf_src)
                if dep != src and dep not in graph[src]:
                    graph[src].append(dep)

    # --- Reports → same pattern as forms ---
    for report in catalog.reports:
        src = _node("report", report.name)
        if report.record_source:
            rs = report.record_source.strip()
            if rs.lower() in {q.lower() for q in known_queries}:
                dep = _node("query", rs)
                if dep not in graph[src]:
                    graph[src].append(dep)
            elif rs.lower() in {t.lower() for t in known_tables}:
                dep = _node("table", rs)
                if dep not in graph[src]:
                    graph[src].append(dep)

    # --- Modules → other objects via DoCmd calls in VBA ---
    for module in catalog.modules:
        if not module.source_code:
            continue
        src = _node("module", module.name)
        for pattern, ref_type in _VBA_OPEN_PATTERNS:
            for match in pattern.finditer(module.source_code):
                ref_name = match.group(1).strip()
                # Validate against known names
                ref_known = False
                if ref_type == "form" and ref_name.lower() in {f.lower() for f in known_forms}:
                    ref_known = True
                elif ref_type == "report" and ref_name.lower() in {r.lower() for r in known_reports}:
                    ref_known = True
                elif ref_type == "query" and ref_name.lower() in {q.lower() for q in known_queries}:
                    ref_known = True
                elif ref_type == "table" and ref_name.lower() in {t.lower() for t in known_tables}:
                    ref_known = True

                if ref_known:
                    dep = _node(ref_type, ref_name)
                    if dep not in graph[src]:
                        graph[src].append(dep)

    return dict(graph)


def find_cycles(graph: dict[str, list[str]]) -> list[list[str]]:
    """
    Find all cycles in the dependency graph using Kahn's algorithm.

    Returns a list of cycle paths (each path is a list of node keys).
    """
    # Compute in-degrees
    in_degree: dict[str, int] = {node: 0 for node in graph}
    for node, deps in graph.items():
        for dep in deps:
            if dep in in_degree:
                in_degree[dep] += 1
            else:
                in_degree[dep] = 1

    # BFS topological sort
    queue = deque(n for n, d in in_degree.items() if d == 0)
    while queue:
        node = queue.popleft()
        for neighbor in graph.get(node, []):
            if neighbor in in_degree:
                in_degree[neighbor] -= 1
                if in_degree[neighbor] == 0:
                    queue.append(neighbor)

    # Nodes still with in_degree > 0 are in cycles
    cycle_nodes = {n for n, d in in_degree.items() if d > 0}
    if not cycle_nodes:
        return []

    # DFS to extract actual cycle paths
    return _extract_cycle_paths(graph, cycle_nodes)


def _extract_cycle_paths(
    graph: dict[str, list[str]],
    cycle_nodes: set[str],
) -> list[list[str]]:
    """Extract the actual cycle paths from the set of nodes known to be in cycles."""
    visited: set[str] = set()
    cycles: list[list[str]] = []

    def dfs(node: str, path: list[str], path_set: set[str]) -> None:
        visited.add(node)
        path.append(node)
        path_set.add(node)

        for neighbor in graph.get(node, []):
            if neighbor not in cycle_nodes:
                continue
            if neighbor not in visited:
                dfs(neighbor, path, path_set)
            elif neighbor in path_set:
                # Found a cycle — extract the cycle portion
                idx = path.index(neighbor)
                cycle = path[idx:] + [neighbor]
                cycles.append(cycle)

        path.pop()
        path_set.discard(node)

    for node in cycle_nodes:
        if node not in visited:
            dfs(node, [], set())

    return cycles


def annotate_catalog_with_dependencies(catalog: AccessCatalog) -> None:
    """
    Build the dependency graph, detect cycles, and annotate the catalog in-place.

    Modifies:
      - catalog.dependency_graph
      - query.in_circular_dependency for affected queries
    """
    graph = build_dependency_graph(catalog)
    catalog.dependency_graph = graph

    cycles = find_cycles(graph)

    if cycles:
        # Collect all nodes that are in any cycle
        cyclic_nodes: set[str] = set()
        for cycle in cycles:
            cyclic_nodes.update(cycle)

        # Mark affected queries
        for query in catalog.queries:
            node = _node("query", query.name)
            if node in cyclic_nodes:
                query.in_circular_dependency = True
