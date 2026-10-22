"""The shared FP&A output folder: the Teams "Validation Results" document library.

Every validation report, Tableau extract, Tableau raw copy and run log is saved
under this one folder, in the same FY / month layout for Actuals and Forecast:

    <shared folder>\\FY27\\2026-10 - October\\
        Forecast Validation\\   (report, Tableau Extracts\\, Tableau Raw Copies\\, Logs\\)
        OPEX\\  Revenue\\  Direct Cost\\  Tableau Raw Copies\\  Logs\\   (Actuals)

Each person adds the Teams folder to their own OneDrive once ("Add shortcut to
My files"), so it has the same name on every PC; the app looks for it in the
current person's OneDrive. Nothing falls back to another person's folder.
``shared_output_root`` in settings.json overrides the search.
"""

from __future__ import annotations

import os
from pathlib import Path

from shared import settings
from shared.friendly_errors import PlainLanguageError

SHARED_FOLDER_NAME = "Financial Planning & Analysis-Validation Results - Documents"
TEAMS_LOCATION = "Teams > Financial Planning & Analysis > Validation Results"


def onedrive_roots() -> list[Path]:
    """The current person's work OneDrive folder(s), most likely first."""
    roots: list[Path] = []
    for variable in ("OneDriveCommercial", "OneDrive"):
        value = os.environ.get(variable)
        if value:
            roots.append(Path(value))
    home = Path(os.environ.get("USERPROFILE") or Path.home())
    try:
        roots += sorted(
            child for child in home.iterdir()
            if child.is_dir() and child.name.casefold().startswith("onedrive - ")
        )
    except OSError:
        pass
    unique: list[Path] = []
    for root in roots:
        if root not in unique and root.is_dir():
            unique.append(root)
    return unique


def _not_available(where: str) -> PlainLanguageError:
    return PlainLanguageError(
        f"The FP&A shared folder is not on this computer ({where}).\n\n"
        f"In {TEAMS_LOCATION}, open Files, choose Open in SharePoint, and click "
        '"Add shortcut to My files". Wait for OneDrive to show the folder, then run again.'
    )


def shared_output_root() -> Path:
    """The synced Teams folder that holds all FP&A automation output."""
    configured = settings.get_path("shared_output_root")
    if configured is not None:
        if not configured.is_dir():
            raise _not_available(f"settings.json points to {configured}, which was not found")
        return configured

    for root in onedrive_roots():
        for candidate in (
            root / "Shortcuts" / SHARED_FOLDER_NAME,
            root / SHARED_FOLDER_NAME,
            root / "My files" / "Shortcuts" / SHARED_FOLDER_NAME,
        ):
            if candidate.is_dir():
                return candidate

    checked = ", ".join(str(root) for root in onedrive_roots()) or "no OneDrive folder was found"
    raise _not_available(f'looked for "{SHARED_FOLDER_NAME}" in {checked}')


def local_work_folder(name: str) -> Path:
    """A folder on this PC only (not synced), for files that exist during a run."""
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / ".local" / "share")
    folder = Path(base) / settings.APP_NAME / name
    folder.mkdir(parents=True, exist_ok=True)
    return folder
