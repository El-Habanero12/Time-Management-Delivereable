from __future__ import annotations

import os
from pathlib import Path

APP_NAME = "HourlyTracker"
PROFILE_ENV_VAR = "HOURLYTRACKER_PROFILE"
_TEST_TOKEN = "TEST"
_DOCS_FOLDER = "Documents"


def is_test_profile() -> bool:
    """Return True when running under the TEST profile."""
    return os.environ.get(PROFILE_ENV_VAR, "").strip().upper() == _TEST_TOKEN


def _profiled_name(base: str) -> str:
    return f"{base}_TEST" if is_test_profile() else base


def get_appdata_dir() -> Path:
    """
    Resolve the base directory under %APPDATA% (or cwd fallback) for app state.
    No directories are created here; callers handle creation.
    """
    appdata_root = os.environ.get("APPDATA")
    base = Path(appdata_root) if appdata_root else Path.cwd()
    return base / _profiled_name(APP_NAME)


def get_docs_dir() -> Path:
    """
    Resolve the user-facing Documents directory for exported/log files.
    No directories are created here; callers handle creation.
    """
    home = Path(os.environ.get("USERPROFILE") or Path.home())
    docs = home / _DOCS_FOLDER
    return docs / _profiled_name(APP_NAME)


def get_default_expenses_path() -> Path:
    """Default location for the user's spending workbook (profile-aware)."""
    return get_docs_dir() / "Expenses.xlsx"
