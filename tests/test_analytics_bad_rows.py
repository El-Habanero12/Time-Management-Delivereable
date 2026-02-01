from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from hourly_tracker.analytics import write_analytics
from hourly_tracker.app import Config


class AnalyticsBadRowTests(unittest.TestCase):
    def test_write_analytics_skips_header_rows_in_task_events(self) -> None:
        base = Path("build_tmp/tests_badrows")
        state = base / "state"
        data = base / "data"
        state.mkdir(parents=True, exist_ok=True)
        data.mkdir(parents=True, exist_ok=True)

        log_path = data / "time_log.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Entries"
        ws.append(["id", "timestamp", "activity", "notes", "category", "energy", "focus", "prompt_type", "start_time", "end_time", "created_at"])
        ws.append(["1", "2026-01-01T10:00:00", "Work", "", "Work", 3, 3, "regular", "", "", "2026-01-01T10:00:00"])

        ws_events = wb.create_sheet("Task_Events")
        ws_events.append(["id", "task_id", "timestamp", "action", "minutes", "effort", "could_be_faster", "notes"])
        ws_events.append(["id", "task_id", "timestamp", "action", "minutes", "effort", "could_be_faster", "notes"])  # duplicated header row that caused crashes
        ws_events.append(["e1", "task_1", "2026-01-01T10:05:00", "worked", "minutes", 3, False, ""])  # bad minutes string
        ws_events.append(["e2", "task_1", "2026-01-01T11:05:00", "worked", 30, 3, False, ""])  # good row
        wb.save(log_path)

        cfg = Config()
        cfg.state_dir = state
        cfg.data_dir = data
        cfg.log_path = log_path
        cfg.report_path = data / "report.html"
        cfg.log_lock_path = state / "time_log.lock"
        cfg.resolve_paths()

        # Should not raise even with malformed rows.
        report_path, _, _ = write_analytics(cfg)
        self.assertTrue(report_path.exists())


if __name__ == "__main__":
    unittest.main()
