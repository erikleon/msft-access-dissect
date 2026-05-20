"""
Rich progress bar wrapper for the extraction pipeline.
"""

from __future__ import annotations

from contextlib import contextmanager
from typing import Generator

from rich.console import Console
from rich.progress import (
    BarColumn,
    MofNCompleteColumn,
    Progress,
    SpinnerColumn,
    TaskProgressColumn,
    TextColumn,
    TimeElapsedColumn,
)

console = Console(stderr=True)


def make_progress() -> Progress:
    return Progress(
        SpinnerColumn(),
        TextColumn("[bold blue]{task.description}"),
        BarColumn(),
        MofNCompleteColumn(),
        TaskProgressColumn(),
        TimeElapsedColumn(),
        console=console,
        transient=False,
    )


@contextmanager
def extraction_progress(total: int, description: str = "Extracting") -> Generator[Progress, None, None]:
    """Context manager that yields a Rich Progress instance pre-configured for extraction."""
    with make_progress() as progress:
        progress.add_task(description, total=total)
        yield progress
