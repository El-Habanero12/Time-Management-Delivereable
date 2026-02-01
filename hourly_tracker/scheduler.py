from __future__ import annotations

import threading
import time
from dataclasses import dataclass
from datetime import date, datetime, time as dt_time, timedelta
from enum import Enum
from typing import Callable, Optional, Any
from .state import SchedulerState, StateStore, apply_pause, apply_snooze


class SchedulerMode(str, Enum):
    RUNNING = "running"
    SNOOZED = "snoozed"
    PAUSED = "paused"
    PROMPTING = "prompting"


class ActionType(str, Enum):
    NONE = "none"
    PROMPT = "prompt"
    CATCH_UP = "catch_up"


@dataclass
class SchedulerAction:
    action: ActionType
    hours_missed: int = 0


def _interval_delta(cfg: Any) -> timedelta:
    return timedelta(minutes=max(1, int(cfg.interval_minutes)))


def _hours_between(start: datetime, end: datetime) -> float:
    return max(0.0, (end - start).total_seconds() / 3600.0)


def compute_scheduler_action(state: SchedulerState, cfg: Any, now: datetime) -> SchedulerAction:
    """Pure decision function suitable for tests."""
    interval = _interval_delta(cfg)

    if state.paused_until and now < state.paused_until:
        return SchedulerAction(ActionType.NONE)

    if state.snoozed_until and now < state.snoozed_until:
        return SchedulerAction(ActionType.NONE)

    last_prompt = state.last_prompt_at
    if not last_prompt:
        return SchedulerAction(ActionType.PROMPT)

    elapsed = now - last_prompt
    if elapsed < interval:
        return SchedulerAction(ActionType.NONE)

    # Detect larger gaps as missed intervals (sleep/lock/reboot).
    hours_elapsed = _hours_between(last_prompt, now)
    interval_hours = max(1.0 / 60.0, cfg.interval_minutes / 60.0)
    missed_intervals = int(hours_elapsed // interval_hours) - 1
    hours_missed = max(0, int(hours_elapsed // 1))

    should_catch_up = missed_intervals >= 1 or hours_elapsed >= (interval_hours * 1.75)
    if should_catch_up:
        if state.last_resume_at and (now - state.last_resume_at) < interval:
            # Already handled a resume recently; prompt normally instead.
            return SchedulerAction(ActionType.PROMPT)
        capped_hours = min(cfg.catch_up_max_hours, max(1, hours_missed))
        return SchedulerAction(ActionType.CATCH_UP, hours_missed=capped_hours)

    return SchedulerAction(ActionType.PROMPT)


class Scheduler:
    def __init__(
        self,
        cfg: Any,
        state_store: StateStore,
        on_prompt: Callable[[], None],
        on_catch_up: Callable[[int], None],
        on_reflection: Optional[Callable[[date], None]] = None,
        tick_seconds: float = 5.0,
    ) -> None:
        self.cfg = cfg
        self.state_store = state_store
        self.on_prompt = on_prompt
        self.on_catch_up = on_catch_up
        self.on_reflection = on_reflection
        self.tick_seconds = tick_seconds

        self._mode = SchedulerMode.RUNNING
        self._stop_event = threading.Event()
        self._thread = threading.Thread(target=self._run_loop, name="hourly-tracker-scheduler", daemon=True)
        self._lock = threading.Lock()

        self.state = self.state_store.load()

    @property
    def mode(self) -> SchedulerMode:
        with self._lock:
            return self._mode

    def start(self) -> None:
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        if self._thread.is_alive():
            self._thread.join(timeout=2)

    def _set_mode(self, mode: SchedulerMode) -> None:
        with self._lock:
            self._mode = mode

    def _run_loop(self) -> None:
        # On startup, make sure workbook scheduling uses persisted state.
        while not self._stop_event.is_set():
            now = datetime.now()
            action = compute_scheduler_action(self.state, self.cfg, now)
            if action.action == ActionType.PROMPT and self.mode != SchedulerMode.PROMPTING:
                self._set_mode(SchedulerMode.PROMPTING)
                try:
                    self.on_prompt()
                finally:
                    self._set_mode(SchedulerMode.RUNNING)
            elif action.action == ActionType.CATCH_UP and self.mode != SchedulerMode.PROMPTING:
                self._set_mode(SchedulerMode.PROMPTING)
                try:
                    self.on_catch_up(action.hours_missed)
                finally:
                    self._set_mode(SchedulerMode.RUNNING)

            if self.on_reflection and self.mode != SchedulerMode.PROMPTING:
                due_date = self._reflection_due_date(now)
                if due_date:
                    self._set_mode(SchedulerMode.PROMPTING)
                    try:
                        self.on_reflection(due_date)
                    finally:
                        self._set_mode(SchedulerMode.RUNNING)

            time.sleep(self.tick_seconds)

    def mark_prompted(self, when: Optional[datetime] = None) -> None:
        now = when or datetime.now()
        self.state.last_prompt_at = now
        self.state.snoozed_until = None
        self.state_store.save(self.state)

    def mark_entry(self, when: Optional[datetime] = None) -> None:
        now = when or datetime.now()
        self.state.last_entry_at = now
        self.state_store.save(self.state)

    def mark_resume_handled(self, when: Optional[datetime] = None) -> None:
        now = when or datetime.now()
        self.state.last_resume_at = now
        self.state_store.save(self.state)

    def snooze(self, minutes: Optional[int] = None) -> None:
        now = datetime.now()
        apply_snooze(self.state, minutes or self.cfg.snooze_minutes, now)
        self._set_mode(SchedulerMode.SNOOZED)
        self.state_store.save(self.state)

    def pause(self, minutes: Optional[int] = None) -> None:
        now = datetime.now()
        apply_pause(self.state, minutes or self.cfg.pause_minutes, now)
        self._set_mode(SchedulerMode.PAUSED)
        self.state_store.save(self.state)

    def resume(self) -> None:
        self.state.paused_until = None
        self.state.snoozed_until = None
        self._set_mode(SchedulerMode.RUNNING)
        self.state_store.save(self.state)

    def mark_reflection_completed(self, day: date) -> None:
        self.state.last_reflection_date = day
        self.state_store.save(self.state)

    def _reflection_due_date(self, now: datetime) -> Optional[date]:
        if not getattr(self.cfg, "reflection_enabled", True):
            return None

        try:
            hh, mm = str(self.cfg.reflection_time_local).split(":")
            target_time = dt_time(hour=int(hh), minute=int(mm))
        except Exception:
            target_time = dt_time(hour=23, minute=30)

        last_done = self.state.last_reflection_date

        today_target = datetime.combine(now.date(), target_time)
        yesterday_target = today_target - timedelta(days=1)

        # Normal case: today, time has passed, and not done today.
        if now >= today_target and last_done != now.date():
            return now.date()

        # Catch-up: missed yesterday (or earlier) and time has passed.
        if now >= yesterday_target and (not last_done or last_done < yesterday_target.date()):
            return yesterday_target.date()

        return None
