from __future__ import annotations

import shutil
from pathlib import Path

from hourly_tracker.paths import (
    get_appdata_dir,
    get_docs_dir,
    get_docs_reflections_dir,
    get_user_expenses_path,
    get_user_time_log_path,
)
from hourly_tracker.resources_util import resource_path


def ensure_user_files_exist() -> dict[str, Path]:
    """
    Ensure per-user working copies exist in the profile's docs folder.
    Returns dict with keys: time_log, expenses, reflections_dir, appdata_dir, docs_dir.
    Idempotent and safe to call multiple times.
    """
    appdata_dir = get_appdata_dir()
    docs_dir = get_docs_dir()
    reflections_dir = get_docs_reflections_dir()

    for p in (appdata_dir, docs_dir, reflections_dir):
        p.mkdir(parents=True, exist_ok=True)

    time_log_path = get_user_time_log_path()
    expenses_path = get_user_expenses_path()

    time_log_template = resource_path("resources/time_log_template.xlsx")
    expenses_template = resource_path("resources/Expenses.xlsx")

    if time_log_template.exists() and not time_log_path.exists():
        shutil.copy2(time_log_template, time_log_path)

    if expenses_template.exists() and not expenses_path.exists():
        shutil.copy2(expenses_template, expenses_path)

    return {
        "appdata_dir": appdata_dir,
        "docs_dir": docs_dir,
        "reflections_dir": reflections_dir,
        "time_log": time_log_path,
        "expenses": expenses_path,
    }
