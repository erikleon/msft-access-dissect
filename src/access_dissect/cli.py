"""
access-dissect CLI

Commands:
  info      Quick summary without full extraction
  extract   Full extraction → catalog JSON/YAML
  render    Render catalog to Markdown/HTML/SQL
  analyze   Dependency + complexity + recommendations from a catalog
  run       Combined extract + render (all formats) + analyze
"""

from __future__ import annotations

import sys
from enum import Enum
from pathlib import Path
from typing import Annotated, Optional

import typer
from rich.console import Console

app = typer.Typer(
    name="access-dissect",
    help="Reverse-engineer, document, and modernize Microsoft Access applications.",
    add_completion=False,
    rich_markup_mode="rich",
)
console = Console()
err_console = Console(stderr=True)


# ---------------------------------------------------------------------------
# Enums for CLI choices
# ---------------------------------------------------------------------------


class ScopeName(str, Enum):
    structure = "structure"
    migration = "migration"
    full = "full"


class FormatName(str, Enum):
    markdown = "markdown"
    html = "html"
    sql = "sql"
    yaml = "yaml"


class DialectName(str, Enum):
    postgres = "postgres"
    sqlite = "sqlite"
    mssql = "mssql"


# ---------------------------------------------------------------------------
# info command
# ---------------------------------------------------------------------------


@app.command("info")
def cmd_info(
    file: Annotated[Path, typer.Argument(help="Path to .accdb or .mdb file")],
    password: Annotated[Optional[str], typer.Option("--password", "-p", help="Database password")] = None,
) -> None:
    """
    [bold]Quick summary[/bold] of an Access database without full extraction.

    Reads database properties and object counts in ~5 seconds.
    """
    from access_dissect.extract.com_session import COMSession, run_preflight_checks
    from access_dissect.extract.properties import extract_properties

    _check_file(file)
    try:
        run_preflight_checks()
    except Exception as e:
        err_console.print(f"[bold red]Preflight check failed:[/bold red] {e}")
        raise typer.Exit(1)

    with COMSession(file, password=password) as session:
        props, _warns = extract_properties(session.app, session.current_db, file)

    console.print()
    console.rule(f"[bold blue]{file.name}")
    console.print(f"  Title:          {props.title or '(none)'}")
    console.print(f"  Author:         {props.author or '(none)'}")
    console.print(f"  Access version: {props.access_version or 'unknown'}")
    console.print(f"  File size:      {(props.file_size_bytes or 0) / 1024:.1f} KB")
    console.print(f"  VBA locked:     {'Yes ⚠' if props.vba_project_locked else 'No'}")
    console.print(f"  Format:         {'ACCDB' if props.is_accdb else 'MDB'}")
    if props.workgroup_db_path:
        console.print(f"  Workgroup MDW:  {props.workgroup_db_path}")
    console.rule()


# ---------------------------------------------------------------------------
# extract command
# ---------------------------------------------------------------------------


@app.command("extract")
def cmd_extract(
    file: Annotated[Path, typer.Argument(help="Path to .accdb or .mdb file")],
    scope: Annotated[ScopeName, typer.Option("--scope", "-s")] = ScopeName.structure,
    password: Annotated[Optional[str], typer.Option("--password", "-p")] = None,
    output: Annotated[Optional[Path], typer.Option("--output", "-o",
        help="Output path for catalog JSON (default: <stem>_catalog.json)")] = None,
    include_system: Annotated[bool, typer.Option("--include-system")] = False,
    max_rows: Annotated[int, typer.Option("--max-rows",
        help="Max rows per table for INSERT scripts (0 = all)")] = 1000,
    timeout: Annotated[float, typer.Option("--timeout",
        help="Per-object COM timeout in seconds")] = 30.0,
    yaml_output: Annotated[bool, typer.Option("--yaml",
        help="Save catalog as YAML instead of JSON")] = False,
) -> None:
    """
    [bold]Extract[/bold] all objects from an Access database to a catalog file.

    The catalog is a machine-readable JSON (or YAML) intermediate format.
    Use [bold]render[/bold] to convert it to Markdown/HTML/SQL.
    """
    from access_dissect.catalog.models import ExtractionScope
    from access_dissect.extract.engine import ExtractionEngine
    from access_dissect.utils.io import save_catalog

    _check_file(file)

    if output is None:
        suffix = ".yaml" if yaml_output else ".json"
        output = file.parent / f"{file.stem}_catalog{suffix}"

    output_dir = output.parent

    engine = ExtractionEngine(
        file_path=file,
        scope=ExtractionScope(scope.value),
        password=password,
        include_system=include_system,
        max_rows=max_rows,
        timeout_per_object=timeout,
        output_dir=output_dir if scope != ScopeName.structure else None,
        verbose=True,
    )

    try:
        catalog = engine.run()
    except Exception as e:
        err_console.print(f"[bold red]Extraction failed:[/bold red] {e}")
        raise typer.Exit(1)

    fmt = "yaml" if yaml_output else "json"
    save_catalog(catalog, output, fmt=fmt)
    console.print(f"\n[green]✓[/green] Catalog saved to: [bold]{output}[/bold]")


# ---------------------------------------------------------------------------
# render command
# ---------------------------------------------------------------------------


@app.command("render")
def cmd_render(
    catalog_path: Annotated[Path, typer.Argument(help="Path to catalog.json or catalog.yaml")],
    format: Annotated[list[FormatName], typer.Option("--format", "-f",
        help="Output format (can repeat: --format markdown --format html)")] = [FormatName.markdown],
    dialect: Annotated[DialectName, typer.Option("--dialect")] = DialectName.postgres,
    output: Annotated[Optional[Path], typer.Option("--output", "-o",
        help="Output path/directory")] = None,
    include_queries: Annotated[bool, typer.Option("--include-queries",
        help="Include translated queries as SQL VIEWs in DDL output")] = False,
) -> None:
    """
    [bold]Render[/bold] a catalog to documentation and migration artifacts.

    Formats: markdown (docs/), html (single file), sql (DDL script), yaml (catalog copy)
    """
    from access_dissect.render.html import render_html
    from access_dissect.render.markdown import render_markdown
    from access_dissect.render.sql import save_ddl
    from access_dissect.utils.io import load_catalog, save_catalog

    if not catalog_path.exists():
        err_console.print(f"[red]Catalog not found: {catalog_path}[/red]")
        raise typer.Exit(1)

    catalog = load_catalog(catalog_path)
    stem = catalog_path.stem.replace("_catalog", "")

    if output is None:
        output = catalog_path.parent

    for fmt in format:
        if fmt == FormatName.markdown:
            md_dir = Path(output) if len(format) > 1 else Path(output) / "docs"
            render_markdown(catalog, md_dir, dialect=dialect.value)
            console.print(f"[green]✓[/green] Markdown docs → [bold]{md_dir}[/bold]")

        elif fmt == FormatName.html:
            html_path = Path(output) / f"{stem}_report.html" if Path(output).is_dir() else Path(output)
            render_html(catalog, html_path)
            console.print(f"[green]✓[/green] HTML report → [bold]{html_path}[/bold]")

        elif fmt == FormatName.sql:
            sql_path = Path(output) / f"{stem}_{dialect.value}.sql" if Path(output).is_dir() else Path(output)
            save_ddl(catalog, sql_path, dialect=dialect.value, include_queries=include_queries)
            console.print(f"[green]✓[/green] SQL DDL → [bold]{sql_path}[/bold]")

        elif fmt == FormatName.yaml:
            yaml_path = Path(output) / f"{stem}_catalog.yaml" if Path(output).is_dir() else Path(output)
            save_catalog(catalog, yaml_path, fmt="yaml")
            console.print(f"[green]✓[/green] YAML catalog → [bold]{yaml_path}[/bold]")


# ---------------------------------------------------------------------------
# analyze command
# ---------------------------------------------------------------------------


@app.command("analyze")
def cmd_analyze(
    catalog_path: Annotated[Path, typer.Argument(help="Path to catalog.json or catalog.yaml")],
    output: Annotated[Optional[Path], typer.Option("--output", "-o")] = None,
) -> None:
    """
    [bold]Analyze[/bold] a catalog: dependency graph, complexity scores, modernization recommendations.

    Produces analysis.md and updates catalog with dependency_graph.
    """
    from access_dissect.analyze.engine import analyze_catalog, render_analysis_markdown
    from access_dissect.utils.io import load_catalog, save_catalog

    if not catalog_path.exists():
        err_console.print(f"[red]Catalog not found: {catalog_path}[/red]")
        raise typer.Exit(1)

    catalog = load_catalog(catalog_path)
    report = analyze_catalog(catalog, verbose=True)

    # Save updated catalog (with dependency_graph populated)
    save_catalog(catalog, catalog_path)

    # Output path for analysis report
    if output is None:
        output = catalog_path.parent / "analysis.md"

    render_analysis_markdown(report, output)
    console.print(f"\n[green]✓[/green] Analysis report → [bold]{output}[/bold]")


# ---------------------------------------------------------------------------
# run command (combined)
# ---------------------------------------------------------------------------


@app.command("run")
def cmd_run(
    file: Annotated[Path, typer.Argument(help="Path to .accdb or .mdb file")],
    scope: Annotated[ScopeName, typer.Option("--scope", "-s")] = ScopeName.structure,
    password: Annotated[Optional[str], typer.Option("--password", "-p")] = None,
    output_dir: Annotated[Optional[Path], typer.Option("--output-dir", "-o")] = None,
    dialect: Annotated[DialectName, typer.Option("--dialect")] = DialectName.postgres,
    timeout: Annotated[float, typer.Option("--timeout")] = 30.0,
    max_rows: Annotated[int, typer.Option("--max-rows")] = 1000,
    no_analyze: Annotated[bool, typer.Option("--no-analyze")] = False,
) -> None:
    """
    [bold]Full pipeline[/bold]: extract + render (all formats) + analyze.

    Produces in output-dir:
      <stem>_catalog.json, docs/ (markdown), <stem>_report.html,
      <stem>_<dialect>.sql, analysis.md
    """
    from access_dissect.analyze.engine import analyze_catalog, render_analysis_markdown
    from access_dissect.catalog.models import ExtractionScope
    from access_dissect.extract.engine import ExtractionEngine
    from access_dissect.render.html import render_html
    from access_dissect.render.markdown import render_markdown
    from access_dissect.render.sql import save_ddl
    from access_dissect.utils.io import save_catalog

    _check_file(file)

    if output_dir is None:
        output_dir = file.parent / f"{file.stem}_dissect"
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    stem = file.stem

    # Extract
    engine = ExtractionEngine(
        file_path=file,
        scope=ExtractionScope(scope.value),
        password=password,
        max_rows=max_rows,
        timeout_per_object=timeout,
        output_dir=output_dir,
        verbose=True,
    )
    try:
        catalog = engine.run()
    except Exception as e:
        err_console.print(f"[bold red]Extraction failed:[/bold red] {e}")
        raise typer.Exit(1)

    # Save catalog
    catalog_path = output_dir / f"{stem}_catalog.json"
    save_catalog(catalog, catalog_path)
    console.print(f"[green]✓[/green] Catalog → [bold]{catalog_path}[/bold]")

    # Render Markdown
    md_dir = output_dir / "docs"
    render_markdown(catalog, md_dir, dialect=dialect.value)
    console.print(f"[green]✓[/green] Markdown docs → [bold]{md_dir}[/bold]")

    # Render HTML
    html_path = output_dir / f"{stem}_report.html"
    render_html(catalog, html_path)
    console.print(f"[green]✓[/green] HTML report → [bold]{html_path}[/bold]")

    # Render SQL DDL
    sql_path = output_dir / f"{stem}_{dialect.value}.sql"
    save_ddl(catalog, sql_path, dialect=dialect.value)
    console.print(f"[green]✓[/green] SQL DDL → [bold]{sql_path}[/bold]")

    # Analyze
    if not no_analyze:
        report = analyze_catalog(catalog, verbose=True)
        save_catalog(catalog, catalog_path)  # save with dependency_graph
        analysis_path = output_dir / "analysis.md"
        render_analysis_markdown(report, analysis_path)
        console.print(f"[green]✓[/green] Analysis → [bold]{analysis_path}[/bold]")

    console.print(f"\n[bold green]✓ Complete![/bold green] Output directory: [bold]{output_dir}[/bold]")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _check_file(path: Path) -> None:
    if not path.exists():
        err_console.print(f"[red]File not found: {path}[/red]")
        raise typer.Exit(1)
    if path.suffix.lower() not in (".accdb", ".mdb"):
        err_console.print(
            f"[yellow]Warning: '{path.name}' does not have .accdb or .mdb extension.[/yellow]"
        )


if __name__ == "__main__":
    app()
