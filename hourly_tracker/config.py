from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict

APP_NAME = "HourlyTracker"


def _default_state_dir() -> Path:
    """Internal state lives under %APPDATA% by default."""
    appdata = os.environ.get("APPDATA")
    if not appdata:
        # Fallback to local repo directory if APPDATA is unavailable.
        return Path.cwd() / APP_NAME
    return Path(appdata) / APP_NAME


def _default_data_dir() -> Path:
    """User-facing files default to Documents\\HourlyTracker."""
    home = Path(os.environ.get("USERPROFILE") or Path.home())
    docs = home / "Documents"
    return docs / APP_NAME


def _default_expenses_path() -> Path:
    """Explicit default for spending log workbook."""
    return Path(r"C:\Users\antho\OneDrive\Desktop\Expenses.xlsx")


def ensure_app_dirs(base_dir: Path) -> Path:
    base_dir.mkdir(parents=True, exist_ok=True)
    return base_dir


@dataclass
class AnalyticsRules:
    # Each check-in represents this many hours unless a gap rule changes it.
    entry_hours: float = 1.0
    # If the gap between entries exceeds this, we break blocks at the gap.
    gap_break_hours: float = 2.0


@dataclass
class Config:
    interval_minutes: int = 60
    snooze_minutes: int = 10
    dismiss_snooze_minutes: int = 10
    pause_minutes: int = 60
    catch_up_max_hours: int = 8

    llm_enabled: bool = False
    llm_model: str = "llama3.1:8b"
    llm_timeout_seconds: int = 20

    no_network_mode: bool = True

    # Directories
    state_dir: Path = field(default_factory=_default_state_dir)
    data_dir: Path = field(default_factory=_default_data_dir)

    # Legacy field kept for backward compatibility; treated as state_dir internally.
    appdata_dir: Path = field(default_factory=_default_state_dir)

    # Paths (resolved during resolve_paths)
    log_path: Path | None = None
    report_path: Path | None = None
    state_path: Path | None = None
    learned_rules_path: Path | None = None
    config_path: Path | None = None
    log_lock_path: Path | None = None
    reflections_dir: Path | None = None
    expenses_path: Path | None = None

    # Reflections
    reflection_enabled: bool = True
    reflection_time_local: str = "23:30"  # HH:MM local time

    # Runtime-only metadata (not persisted across runs but harmless if serialized)
    report_updated_at: str | None = None

    analytics_rules: AnalyticsRules = field(default_factory=AnalyticsRules)

    _PATH_FIELDS = [
        "appdata_dir",
        "state_dir",
        "data_dir",
        "log_path",
        "report_path",
        "state_path",
        "learned_rules_path",
        "config_path",
        "log_lock_path",
        "reflections_dir",
        "expenses_path",
    ]

    def _coerce_path_fields(self) -> None:
        """Ensure every path-like field is a pathlib.Path instance."""
        for name in self._PATH_FIELDS:
            value = getattr(self, name, None)
            if value is None:
                continue
            if not isinstance(value, Path):
                try:
                    setattr(self, name, Path(value))
                except Exception:
                    # Leave as-is if coercion fails; resolve_paths will set defaults.
                    pass

    def resolve_paths(self) -> "Config":
        self._coerce_path_fields()
        # Prefer explicit state/data dirs; fall back to legacy appdata_dir.
        state_base = ensure_app_dirs(Path(self.state_dir or self.appdata_dir))
        data_base = ensure_app_dirs(Path(self.data_dir))

        # Keep legacy field aligned so existing callers continue to work.
        self.state_dir = state_base
        self.appdata_dir = state_base
        self.data_dir = data_base

        if self.log_path is None:
            self.log_path = data_base / "time_log.xlsx"
        if self.report_path is None:
            self.report_path = data_base / "report.html"
        if self.state_path is None:
            self.state_path = state_base / "state.json"
        if self.learned_rules_path is None:
            self.learned_rules_path = state_base / "learned_rules.json"
        if self.config_path is None:
            self.config_path = state_base / "config.json"
        if self.log_lock_path is None:
            self.log_lock_path = state_base / "time_log.lock"
        if self.reflections_dir is None:
            self.reflections_dir = data_base / "reflections"
        if self.expenses_path is None:
            self.expenses_path = _default_expenses_path()
        else:
            self.expenses_path = Path(self.expenses_path)
        return self

    def to_json_dict(self) -> Dict[str, Any]:
        data = asdict(self)
        # Convert Paths to strings for JSON serialization.
        for key in self._PATH_FIELDS:
            value = data.get(key)
            if value is not None:
                data[key] = str(value)
        return data

    @staticmethod
    def _coerce_paths(data: Dict[str, Any]) -> Dict[str, Any]:
        for key in Config._PATH_FIELDS:
            value = data.get(key)
            if value:
                data[key] = Path(value)
        return data


def load_config(path: Path | None = None) -> Config:
    cfg = Config().resolve_paths()
    cfg_path = path or cfg.config_path
    assert cfg_path is not None

    if cfg_path.exists():
        try:
            data = json.loads(cfg_path.read_text(encoding="utf-8"))
            data = Config._coerce_paths(data)
            # Handle nested analytics rules if present.
            analytics_data = data.pop("analytics_rules", None)
            cfg = Config(**data)
            if analytics_data:
                cfg.analytics_rules = AnalyticsRules(**analytics_data)
            cfg.resolve_paths()
        except Exception:
            # If config is corrupt, fall back to defaults but do not overwrite yet.
            cfg = Config().resolve_paths()

    return cfg


def save_config(cfg: Config) -> Path:
    cfg.resolve_paths()
    assert cfg.config_path is not None
    cfg.config_path.write_text(json.dumps(cfg.to_json_dict(), indent=2), encoding="utf-8")
    return cfg.config_path
