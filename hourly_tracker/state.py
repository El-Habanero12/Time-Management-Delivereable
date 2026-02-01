from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Optional


def _parse_dt(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except Exception:
        return None


def _fmt_dt(value: Optional[datetime]) -> Optional[str]:
    if not value:
        return None
    return value.isoformat(timespec="seconds")


def _parse_date(value: Optional[str]) -> Optional[date]:
    if not value:
        return None
    try:
        return date.fromisoformat(value)
    except Exception:
        return None


def _fmt_date(value: Optional[date]) -> Optional[str]:
    if not value:
        return None
    return value.isoformat()


@dataclass
class SchedulerState:
    last_prompt_at: Optional[datetime] = None
    last_entry_at: Optional[datetime] = None
    snoozed_until: Optional[datetime] = None
    paused_until: Optional[datetime] = None
    last_resume_at: Optional[datetime] = None
    last_reflection_date: Optional[date] = None

    def to_json(self) -> dict:
        return {
            "last_prompt_at": _fmt_dt(self.last_prompt_at),
            "last_entry_at": _fmt_dt(self.last_entry_at),
            "snoozed_until": _fmt_dt(self.snoozed_until),
            "paused_until": _fmt_dt(self.paused_until),
            "last_resume_at": _fmt_dt(self.last_resume_at),
            "last_reflection_date": _fmt_date(self.last_reflection_date),
        }

    @staticmethod
    def from_json(data: dict) -> "SchedulerState":
        return SchedulerState(
            last_prompt_at=_parse_dt(data.get("last_prompt_at")),
            last_entry_at=_parse_dt(data.get("last_entry_at")),
            snoozed_until=_parse_dt(data.get("snoozed_until")),
            paused_until=_parse_dt(data.get("paused_until")),
            last_resume_at=_parse_dt(data.get("last_resume_at")),
            last_reflection_date=_parse_date(data.get("last_reflection_date")),
        )


class StateStore:
    def __init__(self, path: Path) -> None:
        # Accept str or Path; normalise so parent/exists calls are safe
        self.path = Path(path) if not isinstance(path, Path) else path
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def load(self) -> SchedulerState:
        if not self.path.exists():
            return SchedulerState()
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            return SchedulerState.from_json(data)
        except Exception:
            return SchedulerState()

    def save(self, state: SchedulerState) -> None:
        payload = state.to_json()
        self.path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def apply_snooze(state: SchedulerState, minutes: int, now: datetime) -> SchedulerState:
    state.snoozed_until = now + timedelta(minutes=minutes)
    return state


def apply_pause(state: SchedulerState, minutes: int, now: datetime) -> SchedulerState:
    state.paused_until = now + timedelta(minutes=minutes)
    state.snoozed_until = None
    return state
