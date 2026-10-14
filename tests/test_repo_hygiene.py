"""Compiled files, backups and Tableau working copies must never be committed."""

import re
import shutil
import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
FORBIDDEN = re.compile(
    r"(^|/)__pycache__/|\.py[cod]$|\.bak$|\.backup$|_backup\.py$|_BEFORE_|(^|/)~\$|\.twb$",
    re.IGNORECASE,
)


def test_no_generated_or_backup_files_are_tracked():
    if not shutil.which("git") or not (ROOT / ".git").exists():
        pytest.skip("not a git checkout")
    tracked = subprocess.run(
        ["git", "ls-files"], cwd=ROOT, capture_output=True, text=True, check=True
    ).stdout.splitlines()
    assert [path for path in tracked if FORBIDDEN.search(path)] == []
