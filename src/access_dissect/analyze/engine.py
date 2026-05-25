"""
AnalysisEngine: orchestrate dependency analysis, complexity scoring, and recommendations.
"""

from __future__ import annotations

from pathlib import Path

from rich.console import Console

from access_dissect.analyze.complexity import score_all_objects
from access_dissect.analyze.dependency import annotate_catalog_with_dependencies
from access_dissect.analyze.recommendations import AnalysisReport, generate_recommendations
from access_dissect.catalog.models import AccessCatalog

console = Console(stderr=True)


def analyze_catalog(catalog: AccessCatalog, verbose: bool = True) -> AnalysisReport:
    """
    Run all analysis passes on a catalog.

    1. Build dependency graph and detect cycles
    2. Score object complexity
    3. Generate modernization recommendations

    Modifies catalog.dependency_graph and query.in_circular_dependency in-place.
    Returns an AnalysisReport with all findings.
    """
    if verbose:
        console.print("[blue]→ Building dependency graph...[/blue]")
    annotate_catalog_with_dependencies(catalog)

    if verbose:
        console.print("[blue]→ Scoring migration complexity...[/blue]")
    scores = score_all_objects(catalog)

    if verbose:
        console.print("[blue]→ Generating modernization recommendations...[/blue]")
    report = generate_recommendations(catalog)

    if verbose:
        _print_analysis_summary(report, scores)

    return report


def render_analysis_markdown(report: AnalysisReport, output_path: Path) -> None:
    """Write the analysis report as a Markdown file."""
    from access_dissect.analyze.complexity import ObjectScore  # noqa

    lines: list[str] = []
    lines.append("# Modernization Analysis Report")
    lines.append("")
    lines.append(f"**Source**: `{report.catalog_file}`")
    lines.append(f"**Effort Tier**: **{report.effort_tier}**")
    lines.append("")
    lines.append("## Summary")
    lines.append("")
    lines.append("| Metric | Value |")
    lines.append("|--------|-------|")
    lines.append(f"| Tables | {report.total_tables} |")
    lines.append(f"| Queries | {report.total_queries} |")
    lines.append(f"| Forms | {report.total_forms} |")
    lines.append(f"| Reports | {report.total_reports} |")
    lines.append(f"| VBA Modules | {report.total_modules} |")
    lines.append(f"| Total VBA Lines | {report.total_vba_lines:,} |")
    lines.append(f"| Complexity (Low/Med/High) | "
                 f"{report.complexity_distribution.get('low', 0)} / "
                 f"{report.complexity_distribution.get('medium', 0)} / "
                 f"{report.complexity_distribution.get('high', 0)} |")
    lines.append("")

    # Group recommendations by category
    categories_order = [
        "MIGRATION_BLOCKERS", "SCHEMA_ISSUES", "QUERY_COMPLEXITY",
        "VBA_MIGRATION", "DATA_QUALITY", "QUICK_WINS",
    ]
    by_category: dict[str, list] = {cat: [] for cat in categories_order}
    for rec in report.recommendations:
        by_category.setdefault(rec.category, []).append(rec)

    for cat in categories_order:
        recs = by_category.get(cat, [])
        if not recs:
            continue
        lines.append(f"## {cat.replace('_', ' ').title()}")
        lines.append("")
        for rec in recs:
            priority_emoji = {
                "blocker": "🔴",
                "high": "🟠",
                "medium": "🟡",
                "low": "🟢",
                "info": "ℹ️",
            }.get(rec.priority, "•")
            lines.append(f"### {priority_emoji} {rec.title}")
            lines.append("")
            lines.append(rec.detail)
            if rec.objects:
                lines.append("")
                lines.append("**Affected objects**:")
                for obj in rec.objects[:15]:
                    lines.append(f"- {obj}")
                if len(rec.objects) > 15:
                    lines.append(f"- _(and {len(rec.objects) - 15} more)_")
            lines.append("")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(lines), encoding="utf-8")


def _print_analysis_summary(report: AnalysisReport, scores: list) -> None:
    console.print()
    console.rule("[bold green]Analysis Complete")
    console.print(f"  Effort tier: [bold cyan]{report.effort_tier}[/bold cyan]")
    dist = report.complexity_distribution
    console.print(
        f"  Complexity: [green]{dist.get('low', 0)} low[/green] / "
        f"[yellow]{dist.get('medium', 0)} medium[/yellow] / "
        f"[red]{dist.get('high', 0)} high[/red]"
    )
    blockers = [r for r in report.recommendations if r.priority == "blocker"]
    if blockers:
        console.print(f"  [bold red]⚠ {len(blockers)} migration blocker(s)[/bold red]")
    console.rule()
