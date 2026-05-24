"""
ExtractionEngine: orchestrates all extractors in the correct order.

Pipeline stages:
  1. Metadata (properties, relations, tables, queries)
  2. VBA (modules via VBIDE)
  3. UI objects (forms, reports — requires DoCmd)
  4. Macros
  5. Data (scope-gated: migration | full)
  6. Data profiling (scope-gated: full only)

Per-object errors are isolated: one bad object records a warning and
the pipeline continues. Only DatabaseOpenError / BitnessError abort early.
"""

from __future__ import annotations

import importlib.metadata
from datetime import datetime
from pathlib import Path
from typing import Any

from rich.console import Console
from rich.progress import (
    BarColumn,
    Progress,
    SpinnerColumn,
    TaskID,
    TextColumn,
    TimeElapsedColumn,
)

from access_dissect.catalog.models import (
    AccessCatalog,
    ExtractionScope,
    ExtractionWarning,
    WarningSeverity,
)
from access_dissect.extract.com_session import COMSession, run_preflight_checks
from access_dissect.extract.data import generate_insert_scripts, update_row_counts
from access_dissect.extract.forms import extract_all_forms
from access_dissect.extract.macros import extract_all_macros
from access_dissect.extract.modules import extract_all_modules
from access_dissect.extract.profiler import profile_all_tables
from access_dissect.extract.properties import extract_properties
from access_dissect.extract.queries import extract_all_queries
from access_dissect.extract.relations import extract_all_relations
from access_dissect.extract.reports import extract_all_reports
from access_dissect.extract.tables import extract_all_tables

console = Console(stderr=True)


def _tool_version() -> str:
    try:
        return f"access-dissect {importlib.metadata.version('access-dissect')}"
    except Exception:
        return "access-dissect (dev)"


class ExtractionEngine:
    """
    Orchestrates the full extraction pipeline.

    Usage:
        engine = ExtractionEngine(
            file_path="path/to/app.accdb",
            scope=ExtractionScope.STRUCTURE,
            password=None,
        )
        catalog = engine.run()
    """

    def __init__(
        self,
        file_path: str | Path,
        scope: ExtractionScope = ExtractionScope.STRUCTURE,
        password: str | None = None,
        include_system: bool = False,
        max_rows: int = 1000,
        timeout_per_object: float = 30.0,
        output_dir: Path | None = None,
        verbose: bool = True,
    ) -> None:
        self.file_path = Path(file_path).resolve()
        self.scope = scope
        self.password = password
        self.include_system = include_system
        self.max_rows = max_rows
        self.timeout_per_object = timeout_per_object
        self.output_dir = output_dir
        self.verbose = verbose

    def run(self) -> AccessCatalog:
        """
        Run the full extraction pipeline.

        Returns a fully-populated AccessCatalog.
        Raises DatabaseOpenError or BitnessError for fatal startup failures.
        """
        # Preflight checks before opening any COM
        run_preflight_checks()

        all_warnings: list[ExtractionWarning] = []

        with Progress(
            SpinnerColumn(),
            TextColumn("[bold blue]{task.description}"),
            BarColumn(),
            TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
            TimeElapsedColumn(),
            console=console,
            transient=False,
            disable=not self.verbose,
        ) as progress:
            task = progress.add_task(
                f"Extracting [cyan]{self.file_path.name}[/cyan]",
                total=9,  # stages
            )

            with COMSession(
                self.file_path,
                password=self.password,
            ) as session:
                app = session.app
                db = session.current_db

                # -------------------------------------------------------
                # Stage 1: Database properties
                # -------------------------------------------------------
                progress.update(task, description="Stage 1/9: Database properties")
                db_props, warns = extract_properties(app, db, self.file_path)
                all_warnings.extend(warns)
                progress.advance(task)

                # -------------------------------------------------------
                # Stage 2: Relations
                # -------------------------------------------------------
                progress.update(task, description="Stage 2/9: Relationships")
                relations, warns = extract_all_relations(db)
                all_warnings.extend(warns)
                progress.advance(task)

                # -------------------------------------------------------
                # Stage 3: Tables
                # -------------------------------------------------------
                progress.update(task, description="Stage 3/9: Tables")
                tables, warns = extract_all_tables(db, include_system=self.include_system)
                all_warnings.extend(warns)
                progress.advance(task)

                # -------------------------------------------------------
                # Stage 4: Queries
                # -------------------------------------------------------
                progress.update(task, description="Stage 4/9: Queries")
                queries, warns = extract_all_queries(db)
                all_warnings.extend(warns)
                progress.advance(task)

                # -------------------------------------------------------
                # Stage 5: VBA modules
                # -------------------------------------------------------
                progress.update(task, description="Stage 5/9: VBA modules")
                modules, warns = extract_all_modules(
                    app,
                    vba_locked=db_props.vba_project_locked,
                )
                all_warnings.extend(warns)
                progress.advance(task)

                # -------------------------------------------------------
                # Stage 6: Forms (requires DoCmd)
                # -------------------------------------------------------
                progress.update(task, description="Stage 6/9: Forms")
                forms, warns = extract_all_forms(
                    app,
                    timeout_seconds=self.timeout_per_object,
                )
                all_warnings.extend(warns)
                progress.advance(task)

                # -------------------------------------------------------
                # Stage 7: Reports (requires DoCmd)
                # -------------------------------------------------------
                progress.update(task, description="Stage 7/9: Reports")
                rpts, warns = extract_all_reports(
                    app,
                    timeout_seconds=self.timeout_per_object,
                )
                all_warnings.extend(warns)
                progress.advance(task)

                # -------------------------------------------------------
                # Stage 8: Macros
                # -------------------------------------------------------
                progress.update(task, description="Stage 8/9: Macros")
                macros, warns = extract_all_macros(app, db)
                all_warnings.extend(warns)
                progress.advance(task)

            # COMSession is now closed — remaining stages use pyodbc

            # -------------------------------------------------------
            # Stage 9: Data (scope-gated)
            # -------------------------------------------------------
            progress.update(task, description="Stage 9/9: Data")

            if self.scope in (ExtractionScope.MIGRATION, ExtractionScope.FULL):
                # Update row counts on table objects
                warns = update_row_counts(self.file_path, tables, password=self.password)
                all_warnings.extend(warns)

                # Generate INSERT scripts
                if self.output_dir:
                    warns = generate_insert_scripts(
                        self.file_path,
                        tables,
                        self.output_dir,
                        max_rows=self.max_rows,
                        password=self.password,
                    )
                    all_warnings.extend(warns)

            table_profiles = []
            if self.scope == ExtractionScope.FULL:
                table_profiles, warns = profile_all_tables(
                    self.file_path,
                    tables,
                    password=self.password,
                )
                all_warnings.extend(warns)

            progress.advance(task)

        # -------------------------------------------------------
        # Assemble catalog
        # -------------------------------------------------------
        catalog = AccessCatalog(
            extracted_at=datetime.now(),
            extracted_by=_tool_version(),
            scope=self.scope,
            properties=db_props,
            tables=tables,
            queries=queries,
            forms=forms,
            reports=rpts,
            macros=macros,
            modules=modules,
            relations=relations,
            table_profiles=table_profiles,
            warnings=all_warnings,
        )

        # Summary
        if self.verbose:
            _print_summary(catalog)

        return catalog


def _print_summary(catalog: AccessCatalog) -> None:
    """Print a post-extraction summary to stderr."""
    console.print()
    console.rule("[bold green]Extraction Complete")
    console.print(
        f"  [bold]Tables[/bold]:    {len(catalog.tables):>4}  "
        f"  [bold]Queries[/bold]:   {len(catalog.queries):>4}"
    )
    console.print(
        f"  [bold]Forms[/bold]:     {len(catalog.forms):>4}  "
        f"  [bold]Reports[/bold]:   {len(catalog.reports):>4}"
    )
    console.print(
        f"  [bold]Modules[/bold]:   {len(catalog.modules):>4}  "
        f"  [bold]VBA lines[/bold]: {catalog.total_vba_lines():>4}"
    )
    console.print(
        f"  [bold]Macros[/bold]:    {len(catalog.macros):>4}  "
        f"  [bold]Relations[/bold]: {len(catalog.relations):>4}"
    )

    error_count = catalog.error_count()
    warn_count = len(catalog.warnings_by_severity(WarningSeverity.WARNING))
    if error_count > 0:
        console.print(
            f"\n  [bold red]⚠ {error_count} error(s)[/bold red], "
            f"[yellow]{warn_count} warning(s)[/yellow] — review catalog.warnings"
        )
    elif warn_count > 0:
        console.print(f"\n  [yellow]{warn_count} warning(s)[/yellow] — review catalog.warnings")
    else:
        console.print("\n  [green]✓ No warnings[/green]")
    console.rule()
