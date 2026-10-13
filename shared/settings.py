"""Application settings shared by every automation.

Settings are layered, later layers win:

1. Built-in defaults below.
2. ``settings.json`` next to the app (the .exe, or the project root when run
   from source). Ship this file with the team's shared-drive paths.
3. ``%APPDATA%\\FPA Automation\\settings.json``: one person's own overrides,
   written by the Settings page in the app.

Path settings changed in the app take effect after a restart.
"""

from __future__ import annotations

import json
import logging
import os
import sys
from pathlib import Path
from typing import Any

APP_NAME = "FPA Automation"

DEFAULTS: dict[str, Any] = {
    # Root folder holding the Tableau workbooks, with OPEX, SAC, Revenue and
    # Direct Cost subfolders. Use the shared-drive path, e.g. \\server\FP&A\Workbooks
    "workbooks_root": r"E:\Workbooks",
    # Actuals validation output. Blank = find the synced OneDrive folder automatically.
    "validation_output_root": "",
    # Forecast validation output. Blank = the synced OneDrive shortcut folder.
    "forecast_validation_output_root": "",
    # Default folders on the Initial Budget Setup page.
    "budget_source_folder": r"E:\Sources\Budgets",
    "budget_output_folder": r"E:\Sources\Forecast",
    # Shows the TEST MODE badge in the header.
    "test_mode": False,
    # DEBUG, INFO, WARNING or ERROR
    "log_level": "INFO",
}

log = logging.getLogger(__name__)


def app_dir() -> Path:
    """Folder containing the .exe when frozen, else the project root."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def user_dir() -> Path:
    base = os.environ.get("APPDATA") or str(Path.home() / ".config")
    return Path(base) / APP_NAME


def shared_settings_file() -> Path:
    return app_dir() / "settings.json"


def user_settings_file() -> Path:
    return user_dir() / "settings.json"


def _read(path: Path) -> dict[str, Any]:
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as exc:
        log.warning("Ignoring unreadable settings file %s: %s", path, exc)
        return {}
    if not isinstance(data, dict):
        log.warning("Ignoring settings file %s: expected a JSON object", path)
        return {}
    return {k: v for k, v in data.items() if k in DEFAULTS}


_cache: dict[str, Any] | None = None


def load(refresh: bool = False) -> dict[str, Any]:
    global _cache
    if _cache is None or refresh:
        merged = dict(DEFAULTS)
        merged.update(_read(shared_settings_file()))
        merged.update(_read(user_settings_file()))
        _cache = merged
    return dict(_cache)


def get(key: str) -> Any:
    return load()[key]


def get_path(key: str) -> Path | None:
    """A path setting, or None when blank."""
    value = str(get(key) or "").strip()
    return Path(os.path.expandvars(value)).expanduser() if value else None


def save_user_settings(values: dict[str, Any]) -> Path:
    """Save only the values that differ from the shared/default settings."""
    baseline = dict(DEFAULTS)
    baseline.update(_read(shared_settings_file()))
    current = _read(user_settings_file())
    current.update({k: v for k, v in values.items() if k in DEFAULTS})
    overrides = {k: v for k, v in current.items() if baseline.get(k) != v}

    path = user_settings_file()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(overrides, indent=2), encoding="utf-8")
    load(refresh=True)
    return path


def workbook_path(*parts: str) -> str:
    """Path under the Tableau workbooks root, e.g. workbook_path("OPEX", "Legal.twb")."""
    return str(Path(str(get("workbooks_root"))).joinpath(*parts))
