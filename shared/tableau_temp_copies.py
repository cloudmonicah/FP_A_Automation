"""Location for temporary Tableau (.twb) working copies.

Temporary copies must NEVER be written next to the source workbooks. They are
placed in a dedicated folder inside the synced Finance validation results
folder (Stacy's OneDrive shortcut) and deleted after each run.
"""
from __future__ import annotations

import time
from pathlib import Path
from uuid import uuid4

TEMP_COPY_FOLDER_NAME = "_Tableau_Temp_Copies"
STALE_AFTER_SECONDS = 12 * 60 * 60


def tableau_temp_folder() -> Path:
    from automations.actuals.validation_output_folders import (
        find_actuals_validation_results_folder,
    )

    folder = Path(find_actuals_validation_results_folder()) / TEMP_COPY_FOLDER_NAME
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def cleanup_stale_temp_copies(folder: Path | None = None) -> None:
    """Remove leftovers from crashed or killed runs (older than 12 hours)."""
    folder = folder or tableau_temp_folder()
    cutoff = time.time() - STALE_AFTER_SECONDS
    for item in folder.glob("*.twb"):
        try:
            if item.stat().st_mtime < cutoff:
                item.unlink()
        except OSError:
            pass


def temp_twb_path(source: str | Path, tag: str) -> Path:
    """Return a unique temp .twb path for ``source`` outside the source folder."""
    source = Path(source)
    folder = tableau_temp_folder()
    cleanup_stale_temp_copies(folder)
    return folder / f"{source.stem}_{tag}_{uuid4().hex[:8]}.twb"
