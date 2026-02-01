from __future__ import annotations

import sys
from pathlib import Path


def resource_path(relative_path: str) -> Path:
    """
    Resolve a resource bundled with the app.
    Works in dev (source tree) and PyInstaller (onefile/onedir).
    """
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)).resolve()
    rel = Path(relative_path)
    if rel.is_absolute():
        return rel
    return base / rel
