from __future__ import annotations

import shutil
from pathlib import Path

from hourly_tracker.paths import get_docs_dir
from hourly_tracker.resources import resource_path


def ensure_user_files_exist() -> tuple[Path, Path]:
    """
    Ensure per-user working copies exist in the profile's docs folder.
    Returns (time_log_path, expenses_path).
    """
    docs_dir = get_docs_dir()
    docs_dir.mkdir(parents=True, exist_ok=True)

    time_log_path = docs_dir / "time_log.xlsx"
    expenses_path = docs_dir / "Expenses.xlsx"
    reflections_dir = docs_dir / "reflections"
    reflections_dir.mkdir(parents=True, exist_ok=True)

    time_log_template = resource_path(Path("resources") / "time_log_template.xlsx")
    expenses_template = resource_path(Path("resources") / "Expenses.xlsx")

    if time_log_template.exists() and not time_log_path.exists():
        shutil.copy2(time_log_template, time_log_path)

    if expenses_template.exists() and not expenses_path.exists():
        shutil.copy2(expenses_template, expenses_path)

    return time_log_path, expenses_path
