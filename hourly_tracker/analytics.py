from __future__ import annotations

import json
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from openpyxl import Workbook

from .excel_store import (
    ENTRIES_SHEET,
    LOOKUP_SHEET,
    atomic_save_workbook,
    file_lock,
    load_or_create_workbook,
    read_entries,
    read_task_events,
    read_tasks,
)
from .llm_ollama import detect_ollama, ollama_narrative_summary

DAILY_SHEET = "Daily_Summaries"
WEEKLY_SHEET = "Weekly_Summaries"
MISSED_SHEET = "Missed_Checkins"
TASK_HISTORY_SHEET = "Task_History"


def _log_info(cfg: Config, message: str) -> None:
    try:
        cfg.resolve_paths()
        log_path = cfg.state_dir / "logs" / "app.log"
        log_path.parent.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().isoformat(timespec="seconds")
        with log_path.open("a", encoding="utf-8") as fh:
            fh.write(f"[{ts}] {message}\n")
    except Exception:
        pass


def _log_warn(cfg: Config, message: str) -> None:
    """Lightweight logger writing to state_dir/app.log without failing analytics."""
    try:
        cfg.resolve_paths()
        log_path = cfg.state_dir / "logs" / "app.log"
        log_path.parent.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().isoformat(timespec="seconds")
        with log_path.open("a", encoding="utf-8") as fh:
            fh.write(f"[{ts}] {message}\n")
    except Exception:
        pass


def _coerce_cfg_paths(cfg: "Config") -> "Config":
    """Ensure any string paths on the config are converted to Path objects."""
    for name in ["state_dir", "data_dir", "log_path", "report_path", "state_path", "learned_rules_path", "config_path", "log_lock_path"]:
        value = getattr(cfg, name, None)
        if value is None:
            continue
        if not isinstance(value, Path):
            try:
                setattr(cfg, name, Path(value))
            except Exception:
                pass
    return cfg


def _to_int(x, default: int = 0) -> int:
    # Defensive conversion so malformed Excel cells (e.g., "minutes") never crash analytics.
    if x is None:
        return default
    if isinstance(x, bool):
        return default
    if isinstance(x, int):
        return x
    if isinstance(x, float):
        return int(x)
    if isinstance(x, str):
        s = x.strip()
        if not s:
            return default
        try:
            return int(float(s))
        except ValueError:
            return default
    return default


def _to_int_minutes(x) -> int:
    """More permissive minutes parser that swallows headers/labels."""
    try:
        if x in (None, "", "minutes"):
            return 0
        if isinstance(x, bool):
            return 0
        if isinstance(x, (int, float)):
            return int(float(x))
        if isinstance(x, str):
            s = x.strip()
            if not s:
                return 0
            return int(float(s))
    except Exception:
        return 0
    return 0


def _parse_dt(value: object) -> Optional[datetime]:
    return parse_timestamp(value)


def parse_timestamp(value: object) -> Optional[datetime]:
    """Accept Excel/ISO strings or datetime; tolerate trailing Z; return None on failure."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    s = str(value).strip()
    if not s:
        return None
    if s.endswith("Z"):
        s = s[:-1]
    # Try full ISO
    try:
        return datetime.fromisoformat(s)
    except Exception:
        pass
    # Try trimming microseconds/seconds
    try:
        parts = s.split(".")[0]
        return datetime.fromisoformat(parts)
    except Exception:
        pass
    # Try space separated date time
    try:
        return datetime.fromisoformat(s.replace(" ", "T"))
    except Exception:
        pass
    return None


@dataclass
class Block:
    start: datetime
    end: datetime
    category: str
    activity: str
    energy: int
    focus: int

    @property
    def hours(self) -> float:
        return max(0.0, (self.end - self.start).total_seconds() / 3600.0)


def entries_to_blocks(entries: List[dict], cfg: Config) -> Tuple[List[Block], int]:
    """Convert check-ins to estimated time blocks.

    Heuristic: each entry represents the preceding entry_hours unless a large gap
    suggests a shorter block ending at the next check-in.
    """
    rules = cfg.analytics_rules
    entry_hours = max(0.25, float(rules.entry_hours))
    gap_break_hours = max(entry_hours, float(rules.gap_break_hours))

    parsed: List[Tuple[datetime, dict]] = []
    parse_failures = 0
    for row in entries:
        ts = parse_timestamp(row.get("timestamp"))
        if ts:
            parsed.append((ts, row))
        else:
            parse_failures += 1
    parsed.sort(key=lambda x: x[0])

    blocks: List[Block] = []
    for idx, (ts, row) in enumerate(parsed):
        next_ts = parsed[idx + 1][0] if idx + 1 < len(parsed) else None
        category = str(row.get("category") or "Other")
        activity = str(row.get("activity") or "")
        energy = _to_int(row.get("energy"), 3)
        focus = _to_int(row.get("focus"), 3)

        start = ts - timedelta(hours=entry_hours)
        end = ts

        if next_ts:
            gap_hours = max(0.0, (next_ts - ts).total_seconds() / 3600.0)
            if gap_hours >= gap_break_hours:
                end = ts
            else:
                # Cap the block to avoid overlapping too far into the next entry.
                end = min(ts, next_ts)

        if end <= start:
            start = end - timedelta(hours=entry_hours)

        blocks.append(Block(start=start, end=end, category=category, activity=activity, energy=energy, focus=focus))

    return blocks, parse_failures


def _group_blocks_by_day(blocks: Iterable[Block]) -> Dict[date, List[Block]]:
    grouped: Dict[date, List[Block]] = defaultdict(list)
    for block in blocks:
        grouped[block.end.date()].append(block)
    return dict(grouped)


def _group_blocks_by_week(blocks: Iterable[Block]) -> Dict[date, List[Block]]:
    grouped: Dict[date, List[Block]] = defaultdict(list)
    for block in blocks:
        week_start = block.end.date() - timedelta(days=block.end.weekday())
        grouped[week_start].append(block)
    return dict(grouped)


def _hours_per_category(blocks: Iterable[Block]) -> Dict[str, float]:
    totals: Dict[str, float] = defaultdict(float)
    for block in blocks:
        totals[block.category] += block.hours
    return dict(totals)


def _top_categories(cat_hours: Dict[str, float], n: int = 3) -> List[Tuple[str, float]]:
    return sorted(cat_hours.items(), key=lambda kv: kv[1], reverse=True)[:n]


def _most_common_activities(blocks: Iterable[Block], n: int = 5) -> List[Tuple[str, int]]:
    counter = Counter(block.activity.strip() for block in blocks if block.activity.strip())
    return counter.most_common(n)


def _time_sinks(cat_hours: Dict[str, float]) -> List[Tuple[str, float]]:
    return sorted(cat_hours.items(), key=lambda kv: kv[1], reverse=True)[:5]


def _heuristic_narrative(day_blocks: List[Block], cat_hours: Dict[str, float]) -> Tuple[str, List[str]]:
    top = _top_categories(cat_hours, n=3)
    if not top:
        return ("No check-ins recorded.", ["Set a shorter interval to build the habit.", "Use 'Log now' after long breaks.", "Add categories to the Lookup sheet."])

    top_text = ", ".join(f"{cat} ({hours:.1f}h)" for cat, hours in top)
    avg_focus = sum(b.focus for b in day_blocks) / max(1, len(day_blocks))
    avg_energy = sum(b.energy for b in day_blocks) / max(1, len(day_blocks))

    narrative = (
        f"Most of the day went to {top_text}. "
        f"Average focus was {avg_focus:.1f}/5 and energy was {avg_energy:.1f}/5."
    )
    suggestions = [
        "Protect a 60-90 minute block for your top priority.",
        "Batch admin/email into one or two slots.",
        "If energy is low, schedule a short recovery break.",
    ]
    return narrative, suggestions


def _maybe_llm_summary(cfg: Config, summary_input: str) -> Optional[Tuple[str, List[str]]]:
    if not cfg.llm_enabled:
        return None
    status = detect_ollama()
    if not status.installed:
        return None
    if cfg.llm_model not in status.models:
        return None

    result = ollama_narrative_summary(cfg.llm_model, cfg.llm_timeout_seconds, summary_input)
    if not result:
        return None
    narrative = str(result.get("narrative") or "").strip()
    suggestions_raw = result.get("suggestions") or []
    suggestions = [str(s).strip() for s in suggestions_raw if str(s).strip()]
    if not narrative:
        return None
    if len(suggestions) < 3:
        suggestions = (suggestions + ["Review your top category and trim one low-value task."])[:3]
    return narrative, suggestions[:3]


def _build_summary_input(day: date, blocks: List[Block], cat_hours: Dict[str, float]) -> str:
    parts = [f"Date: {day.isoformat()}"]
    parts.append("Category hours:")
    for cat, hours in sorted(cat_hours.items(), key=lambda kv: kv[1], reverse=True):
        parts.append(f"- {cat}: {hours:.2f}h")
    parts.append("Activities:")
    for b in blocks[:20]:
        parts.append(f"- {b.end.strftime('%H:%M')} {b.category}: {b.activity}")
    return "\n".join(parts)


def _estimate_missed_checkins(entries: List[dict], cfg: Config) -> List[dict]:
    interval_hours = max(1.0 / 60.0, cfg.interval_minutes / 60.0)
    gap_break_hours = max(interval_hours, float(cfg.analytics_rules.gap_break_hours))

    parsed: List[datetime] = []
    for row in entries:
        ts = _parse_dt(row.get("timestamp"))
        if ts:
            parsed.append(ts)
    parsed.sort()
    if len(parsed) < 2:
        return []

    by_day: Dict[date, List[datetime]] = defaultdict(list)
    for ts in parsed:
        by_day[ts.date()].append(ts)

    reports: List[dict] = []
    for day, stamps in sorted(by_day.items()):
        stamps.sort()
        span_hours = max(1.0, (stamps[-1] - stamps[0]).total_seconds() / 3600.0)
        expected = int(span_hours // interval_hours) + 1
        actual = len(stamps)

        largest_gap = 0.0
        for a, b in zip(stamps, stamps[1:]):
            gap = (b - a).total_seconds() / 3600.0
            largest_gap = max(largest_gap, gap)

        missed = max(0, expected - actual)
        if largest_gap >= gap_break_hours or missed > 0:
            reports.append(
                {
                    "date": day.isoformat(),
                    "expected": expected,
                    "actual": actual,
                    "missed": missed,
                    "largest_gap_hours": round(largest_gap, 2),
                }
            )

    return reports


def _upsert_sheet(wb: Workbook, name: str, headers: List[str], rows: List[List[object]]) -> None:
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(title=name)
    ws.append(headers)
    for row in rows:
        ws.append(row)


def write_analytics(cfg: Config) -> tuple[Path, datetime, List[dict]]:
    if not getattr(cfg, "analytics_enabled", True):
        now = datetime.now()
        fallback_path = cfg.report_path or (cfg.data_dir / "report.html")
        return fallback_path, now, []

    cfg = _coerce_cfg_paths(cfg)
    cfg.resolve_paths()
    assert cfg.log_path is not None

    updated_at = datetime.now()

    warnings: List[str] = []

    report_path = cfg.report_path or (cfg.data_dir / "report.html")

    def _safe_read(label: str, fn, fallback):
        try:
            return fn()
        except (TimeoutError, PermissionError) as exc:
            msg = f"{label} unavailable: {exc}"
            warnings.append(msg)
            _log_warn(cfg, msg)
            return fallback
        except Exception as exc:
            msg = f"{label} failed: {exc}"
            warnings.append(msg)
            _log_warn(cfg, msg)
            return fallback

    entries = _safe_read("entries", lambda: read_entries(cfg.log_path, lock_path=cfg.log_lock_path), [])
    total_entries = len(entries)
    detected_headers = list(entries[0].keys()) if entries else list(getattr(read_entries, "last_headers", []))
    if not detected_headers:
        _log_warn(cfg, "Entries header row not found; analytics skipped.")
        write_html_report(
            report_path,
            [],
            [],
            [],
            [],
            [],
            updated_at,
            log_path=str(cfg.log_path),
            total_entries=total_entries,
            today_entries=0,
            warnings=warnings,
        )
        cfg.report_updated_at = updated_at.isoformat(timespec="seconds")
        return report_path, updated_at, entries
    _log_info(
        cfg,
        f"write_analytics log_path={Path(cfg.log_path).resolve()} report_path={report_path.resolve()} entries={total_entries}",
    )
    _log_info(cfg, f"entries headers={detected_headers}")
    task_events = _safe_read("task events", lambda: read_task_events(cfg.log_path, lock_path=cfg.log_lock_path), [])
    tasks = _safe_read("tasks", lambda: read_tasks(cfg.log_path, status_filter=None, lock_path=cfg.log_lock_path), [])
    blocks, parse_failures = entries_to_blocks(entries, cfg)

    daily_blocks = _group_blocks_by_day(blocks)
    weekly_blocks = _group_blocks_by_week(blocks)
    missed_reports = _estimate_missed_checkins(entries, cfg)

    daily_rows: List[List[object]] = []
    today = datetime.now().date()
    today_entry_count = 0
    for day, day_list in sorted(daily_blocks.items()):
        cat_hours = _hours_per_category(day_list)
        top = _top_categories(cat_hours)
        top_text = ", ".join(f"{c}:{h:.1f}" for c, h in top)
        summary_input = _build_summary_input(day, day_list, cat_hours)
        llm_summary = _maybe_llm_summary(cfg, summary_input)
        if llm_summary:
            narrative, suggestions = llm_summary
        else:
            narrative, suggestions = _heuristic_narrative(day_list, cat_hours)

        if day == today:
            today_entry_count = len(day_list)

        daily_rows.append(
            [
                day.isoformat(),
                round(sum(cat_hours.values()), 2),
                top_text,
                json.dumps(cat_hours),
                narrative,
                json.dumps(suggestions),
            ]
        )

    weekly_rows: List[List[object]] = []
    for week_start, week_list in sorted(weekly_blocks.items()):
        week_end = week_start + timedelta(days=6)
        cat_hours = _hours_per_category(week_list)
        top = _top_categories(cat_hours, n=5)
        top_text = ", ".join(f"{c}:{h:.1f}" for c, h in top)
        frequent = _most_common_activities(week_list, n=7)
        frequent_text = ", ".join(f"{a} ({n})" for a, n in frequent)
        sinks = _time_sinks(cat_hours)
        sinks_text = ", ".join(f"{c}:{h:.1f}" for c, h in sinks)

        summary_input = (
            f"Week: {week_start.isoformat()} to {week_end.isoformat()}\n"
            f"Category hours: {cat_hours}\n"
            f"Frequent activities: {frequent_text}\n"
            f"Time sinks: {sinks_text}\n"
        )
        llm_summary = _maybe_llm_summary(cfg, summary_input)
        if llm_summary:
            narrative, suggestions = llm_summary
        else:
            narrative = f"Top categories: {top_text}. Frequent activities: {frequent_text}."
            suggestions = [
                "Cut one recurring low-value activity by 30 minutes.",
                "Schedule deep work early in the week.",
                "Review your category mix on Sunday night.",
            ]

        pie_data = [{"category": c, "hours": round(h, 2)} for c, h in sorted(cat_hours.items(), key=lambda kv: kv[1], reverse=True)]
        weekly_rows.append(
            [
                week_start.isoformat(),
                week_end.isoformat(),
                round(sum(cat_hours.values()), 2),
                top_text,
                frequent_text,
                sinks_text,
                json.dumps(pie_data),
                narrative,
                json.dumps(suggestions),
            ]
        )

    missed_rows = [
        [r["date"], r["expected"], r["actual"], r["missed"], r["largest_gap_hours"]] for r in missed_reports
    ]

    task_history_rows: List[List[object]] = []
    for ev in task_events:
        task_history_rows.append(
            [
                ev.get("timestamp"),
                ev.get("task_id"),
                ev.get("action"),
                ev.get("minutes"),
                ev.get("effort"),
                ev.get("could_be_faster"),
                ev.get("notes"),
            ]
        )

    if parse_failures:
        warnings.append(f"Warning: {parse_failures} entries had unparseable timestamps and were skipped.")
        _log_warn(cfg, warnings[-1])
    try:
        write_html_report(
            report_path,
            daily_rows,
            weekly_rows,
            missed_rows,
            task_history_rows,
            tasks,
            updated_at,
            log_path=str(Path(cfg.log_path).resolve()),
            total_entries=total_entries,
            today_entries=today_entry_count,
            warnings=warnings,
        )
    except Exception as exc:
        _log_warn(cfg, f"write_html_report failed: {exc}")
    cfg.report_updated_at = updated_at.isoformat(timespec="seconds")
    lock_path = cfg.log_lock_path or cfg.log_path.with_suffix(cfg.log_path.suffix + ".lock")
    try:
        with file_lock(lock_path, timeout_seconds=5.0):
            wb = load_or_create_workbook(cfg.log_path)
            _upsert_sheet(
                wb,
                DAILY_SHEET,
                ["date", "total_hours", "top_categories", "category_breakdown_json", "narrative", "suggestions_json"],
                daily_rows,
            )
            _upsert_sheet(
                wb,
                WEEKLY_SHEET,
                [
                    "week_start",
                    "week_end",
                    "total_hours",
                    "top_categories",
                    "frequent_activities",
                    "time_sinks",
                    "pie_data_json",
                    "narrative",
                    "suggestions_json",
                ],
                weekly_rows,
            )
            _upsert_sheet(
                wb,
                MISSED_SHEET,
                ["date", "expected", "actual", "missed", "largest_gap_hours"],
                missed_rows,
            )
            _upsert_sheet(
                wb,
                TASK_HISTORY_SHEET,
                ["timestamp", "task_id", "action", "minutes", "effort", "could_be_faster", "notes"],
                task_history_rows,
            )
            atomic_save_workbook(wb, cfg.log_path)
    except (TimeoutError, PermissionError) as exc:
        _log_warn(cfg, f"Skipped Excel analytics write: {exc}")
    except Exception as exc:
        # If the workbook is open/locked, keep the HTML report updated anyway.
        _log_warn(cfg, f"Analytics workbook update failed: {exc}")
    for msg in warnings:
        _log_warn(cfg, msg)
    return report_path, updated_at, entries


def _daily_task_summary(task_events: List[dict]) -> Tuple[List[List[object]], bool]:
    by_day: Dict[str, Dict[str, int]] = defaultdict(dict)
    invalid_row_seen = False
    for ev in task_events:
        ts = ev.get("timestamp")
        task_id = ev.get("task_id")
        raw = ev.get("minutes")
        if isinstance(raw, str):
            try:
                float(raw.strip())
            except Exception:
                invalid_row_seen = True
                continue
        minutes = _to_int_minutes(raw)
        if minutes == 0 and raw not in (None, "", 0) and str(raw).strip().lower() not in ("0", "minutes", "minute"):
            invalid_row_seen = True
            continue  # skip bad row safely
        if not ts or not task_id:
            continue
        day = str(ts)[:10]
        day_map = by_day.setdefault(day, {})
        day_map[str(task_id)] = day_map.get(str(task_id), 0) + minutes

    rows: List[List[object]] = []
    for day, task_map in sorted(by_day.items()):
        top = sorted(task_map.items(), key=lambda kv: kv[1], reverse=True)[:5]
        rows.append([day, json.dumps(top), sum(task_map.values())])
    return rows, invalid_row_seen


def write_html_report(
    report_path: Path,
    daily_rows: List[List[object]],
    weekly_rows: List[List[object]],
    missed_rows: List[List[object]],
    task_history_rows: List[List[object]],
    tasks: List[dict],
    updated_at: datetime,
    log_path: str,
    total_entries: int,
    today_entries: int,
    warnings: List[str],
) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    updated_text = updated_at.isoformat(timespec="seconds")

    def _rows_to_table(headers: List[str], rows: List[List[object]]) -> str:
        head_html = "".join(f"<th>{h}</th>" for h in headers)
        body_parts: List[str] = []
        for row in rows:
            cells = "".join(f"<td>{str(c)}</td>" for c in row)
            body_parts.append(f"<tr>{cells}</tr>")
        body_html = "\n".join(body_parts) if body_parts else "<tr><td colspan='99'>No data</td></tr>"
        return f"<table><thead><tr>{head_html}</tr></thead><tbody>{body_html}</tbody></table>"

    daily_table = _rows_to_table(
        ["date", "total_hours", "top_categories", "narrative", "suggestions_json"],
        [[r[0], r[1], r[2], r[4], r[5]] for r in daily_rows[-14:]],
    )
    weekly_table = _rows_to_table(
        ["week_start", "week_end", "total_hours", "top_categories", "narrative", "suggestions_json"],
        [[r[0], r[1], r[2], r[3], r[7], r[8]] for r in weekly_rows[-8:]],
    )
    missed_table = _rows_to_table(
        ["date", "expected", "actual", "missed", "largest_gap_hours"],
        missed_rows[-30:],
    )

    task_title_by_id = {str(t.get("id")): str(t.get("title") or "") for t in tasks}
    task_history_display = []
    for row in task_history_rows[-100:]:
        task_id = str(row[1])
        title = task_title_by_id.get(task_id, "")
        task_history_display.append([row[0], f"{task_id} | {title}".strip(" |"), row[2], row[3], row[4], row[5], row[6]])

    task_history_table = _rows_to_table(
        ["timestamp", "task", "action", "minutes", "effort", "could_be_faster", "notes"],
        task_history_display,
    )

    daily_task_rows, invalid_task_history = _daily_task_summary(
        [
            {
                "timestamp": r[0],
                "task_id": r[1],
                "minutes": r[3],
            }
            for r in task_history_rows
        ]
    )
    daily_task_rows = [
        [
            r[0],
            json.dumps([(task_title_by_id.get(t[0], t[0]), t[1]) for t in json.loads(r[1])]),
            r[2],
        ]
        for r in daily_task_rows
    ]
    if invalid_task_history:
        warnings.append("Skipped invalid task rows in Task_Events (non-numeric minutes).")
    daily_task_table = _rows_to_table(
        ["date", "top_tasks_json", "total_minutes"],
        daily_task_rows[-30:],
    )

    html = f"""
<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\" />
  <title>Hourly Tracker Report</title>
  <style>
    body {{ font-family: Segoe UI, Arial, sans-serif; margin: 24px; color: #111; }}
    h1, h2 {{ margin-bottom: 8px; }}
    table {{ border-collapse: collapse; width: 100%; margin: 12px 0 24px; }}
    th, td {{ border: 1px solid #ddd; padding: 8px; font-size: 13px; vertical-align: top; }}
    th {{ background: #f4f6f8; text-align: left; position: sticky; top: 0; }}
    .muted {{ color: #555; font-size: 12px; }}
    code {{ background: #f2f2f2; padding: 2px 4px; border-radius: 4px; }}
  </style>
</head>
<body>
  <h1>Hourly Tracker Report</h1>
  <p class=\"muted\">Report updated at {updated_text} (local time).</p>
  <p class=\"muted\">
    Source workbook: {log_path}<br/>
    Total entries loaded: {total_entries}<br/>
    Entries today: {today_entries}<br/>
    {(' '.join(warnings)) if warnings else ''}
  </p>

  <h2>Daily Summaries</h2>
  {daily_table}

  <h2>Weekly Summaries</h2>
  {weekly_table}

  <h2>Missed Check-Ins</h2>
  {missed_table}

  <h2>Daily Task Summary</h2>
  {daily_task_table}

  <h2>Task History (Recent)</h2>
  {task_history_table}

  <p class=\"muted\">All analytics are computed locally. No network calls are required.</p>
</body>
</html>
"""
    report_path.write_text(html.strip(), encoding="utf-8")
