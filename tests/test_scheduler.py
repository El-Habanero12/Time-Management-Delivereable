from __future__ import annotations

import unittest
from datetime import date, datetime, timedelta
from pathlib import Path

from hourly_tracker.app import Config
from hourly_tracker.scheduler import ActionType, Scheduler, compute_scheduler_action
from hourly_tracker.state import SchedulerState, StateStore


class SchedulerDecisionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.cfg = Config(interval_minutes=60, snooze_minutes=10, pause_minutes=60, catch_up_max_hours=8)

    def test_first_run_prompts_immediately(self) -> None:
        state = SchedulerState(last_prompt_at=None)
        action = compute_scheduler_action(state, self.cfg, now=datetime(2026, 1, 27, 10, 0, 0))
        self.assertEqual(action.action, ActionType.PROMPT)

    def test_snoozed_blocks_prompt(self) -> None:
        now = datetime(2026, 1, 27, 10, 0, 0)
        state = SchedulerState(last_prompt_at=now - timedelta(hours=2), snoozed_until=now + timedelta(minutes=5))
        action = compute_scheduler_action(state, self.cfg, now=now)
        self.assertEqual(action.action, ActionType.NONE)

    def test_paused_blocks_prompt(self) -> None:
        now = datetime(2026, 1, 27, 10, 0, 0)
        state = SchedulerState(last_prompt_at=now - timedelta(hours=2), paused_until=now + timedelta(minutes=30))
        action = compute_scheduler_action(state, self.cfg, now=now)
        self.assertEqual(action.action, ActionType.NONE)

    def test_regular_prompt_after_interval(self) -> None:
        now = datetime(2026, 1, 27, 12, 0, 0)
        state = SchedulerState(last_prompt_at=now - timedelta(hours=1, minutes=1))
        action = compute_scheduler_action(state, self.cfg, now=now)
        self.assertEqual(action.action, ActionType.PROMPT)

    def test_catch_up_after_large_gap(self) -> None:
        now = datetime(2026, 1, 27, 18, 0, 0)
        state = SchedulerState(last_prompt_at=now - timedelta(hours=5))
        action = compute_scheduler_action(state, self.cfg, now=now)
        self.assertEqual(action.action, ActionType.CATCH_UP)
        self.assertGreaterEqual(action.hours_missed, 1)

    def test_resume_recently_handles_catchup_once(self) -> None:
        now = datetime(2026, 1, 27, 18, 0, 0)
        state = SchedulerState(
            last_prompt_at=now - timedelta(hours=5),
            last_resume_at=now - timedelta(minutes=30),
        )
        action = compute_scheduler_action(state, self.cfg, now=now)
        self.assertEqual(action.action, ActionType.PROMPT)


class ReflectionSchedulingTests(unittest.TestCase):
    def setUp(self) -> None:
        base = Path.cwd() / "build_tmp" / "tests_reflection"
        base.mkdir(parents=True, exist_ok=True)
        (base / "state").mkdir(exist_ok=True)
        (base / "data").mkdir(exist_ok=True)

        self.cfg = Config(reflection_enabled=True, reflection_time_local="23:30")
        self.cfg.state_dir = base / "state"
        self.cfg.data_dir = base / "data"
        self.cfg.resolve_paths()
        self.state_store = StateStore(self.cfg.state_path)
        self.scheduler = Scheduler(
            cfg=self.cfg,
            state_store=self.state_store,
            on_prompt=lambda: None,
            on_catch_up=lambda _: None,
            on_reflection=None,
            tick_seconds=0.1,
        )

    def test_reflection_catch_up_for_previous_day(self) -> None:
        now = datetime(2026, 1, 28, 8, 0, 0)
        due = self.scheduler._reflection_due_date(now)
        self.assertEqual(due, date(2026, 1, 27))

    def test_reflection_not_repeated_same_day(self) -> None:
        self.scheduler.state.last_reflection_date = date(2026, 1, 27)
        now = datetime(2026, 1, 27, 23, 40, 0)
        due = self.scheduler._reflection_due_date(now)
        self.assertIsNone(due)


if __name__ == "__main__":
    unittest.main()
