from __future__ import annotations

import sys
from pathlib import Path


def resource_path(relative: str | Path) -> Path:
    """
    Resolve a resource bundled with the app.
    Works in dev (source tree) and PyInstaller (onefile/onedir).
    """
    base = getattr(sys, "_MEIPASS", None)
    root = Path(base) if base else Path(__file__).resolve().parent
    return root / Path(relative)
