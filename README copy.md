# HourlyTracker (Local-Only, Windows 11)

A local-first desktop time tracker that prompts every 60 minutes: **"What are you doing right now?"** Data is stored in a stable Excel `.xlsx` workbook with offline analytics and optional local LLM summaries via Ollama.

This project is designed to run fully offline on Windows 11.

## What You Get
- System tray app with menu actions: Log now, Snooze 10m, Pause 1h, Open log, Open report, Log today's spending, Create shortcuts, Quit.
- Hourly prompts with an always-on-top dialog.
- Fields: timestamp, activity, optional notes, category, energy (1?5), focus (1?5).
- Task manager prompts: add new todos, log work/completions, effort, and reflection.
- Tray Tasks Manager: view/edit/complete tasks outside the hourly prompt.
- Desktop shortcuts helper (data folder, log, report) created once on first launch or via tray menu.
- Daily spending capture via tray into `Expenses.xlsx` on Desktop.
- Nightly reflection prompt (23:30 local default) writes one docx per day under `reflections/`.
- Excel persistence via `openpyxl` only (no pandas).
- Atomic writes and a lock file to reduce corruption risk.
- Local analytics: daily summaries, weekly summaries, missed check-ins, and an HTML report.
- Smart tagging: rule-based suggestion plus learned overrides stored locally.
- Optional local LLM mode (Ollama) with safe fallback to heuristics.
- A no-network guardrail enabled by default.

## Architecture
Key design goals: offline-first, robust scheduling, safe local persistence, and transparent analytics.

### Components
- Tray UI: `pystray` icon and menu.
- Prompt UI: `tkinter` dialogs run on a dedicated Tk thread.
- Scheduler: background thread with a small state machine and persisted state.
- Storage: Excel workbook with stable headers plus a Lookup sheet for categories.
- Analytics: local heuristics that produce daily/weekly summaries and an HTML report.
- Optional local LLM: Ollama CLI integration only when enabled and available.
- Safety: lock file + atomic replace, plus no-network guardrail.

### Component Diagram
```text
+------------------+        +------------------+
|  System Tray UI  |------->|   Scheduler      |
|  (pystray)       |        |  (threaded)      |
+---------+--------+        +--------+---------+
          |                          |
          | menu actions             | due actions
          v                          v
+---------+--------+        +--------+---------+
|  Dialog Runner   |<------>|  State Store     |
|  (Tk thread)     |        |  (state.json)    |
+---------+--------+        +--------+---------+
          |                          |
          | entries                  |
          v                          v
+---------+------------------------------------+
|               Excel Store                    |
| time_log.xlsx: Entries + Lookup + Summaries  |
| lock file + atomic writes                    |
+---------+----------------+-------------------+
          |                |
          | read           | write
          v                v
+---------+--------+  +----+--------------------+
| Tagging Engine   |  | Analytics + Formatting  |
| rules + learning |  | summaries + charts      |
+------------------+  +-------------------------+
```

## Data Schema (Excel)
The workbook uses stable column names. New columns should be appended rather than renamed.

### Sheet: `Entries`

| Column      | Type    | Description |
|-------------|---------|-------------|
| id          | string  | Unique entry id (timestamp-based) |
| timestamp   | string  | ISO timestamp when prompted/logged |
| activity    | string  | Free-text activity |
| notes       | string  | Optional notes |
| category    | string  | Category label |
| energy      | int     | 1?5 |
| focus       | int     | 1?5 |
| prompt_type | string  | regular, manual, catch_up, driving |
| start_time  | string  | Optional ISO start time |
| end_time    | string  | Optional ISO end time |
| created_at  | string  | ISO time of persistence |

### Sheet: `Lookup`

| Column    | Type   | Description |
|-----------|--------|-------------|
| category  | string | Category name |
| keywords  | string | Optional keyword hints |
| regex     | string | Optional regex hints |
| updated_at| string | ISO timestamp |

### Sheet: `Tasks`
| Column         | Type   | Description |
|----------------|--------|-------------|
| id             | string | Task id |
| title          | string | Task title |
| status         | string | open or done |
| created_at     | string | ISO timestamp |
| completed_at   | string | ISO timestamp |
| last_worked_at | string | ISO timestamp |
| total_minutes  | int    | Accumulated minutes worked |
| notes          | string | Optional notes |

### Sheet: `Task_Events`
| Column          | Type   | Description |
|-----------------|--------|-------------|
| id              | string | Event id |
| task_id         | string | Task id |
| timestamp       | string | When the action was logged |
| action          | string | worked or completed |
| minutes         | int    | Minutes spent |
| effort          | int    | 1–5 |
| could_be_faster | bool   | Reflection flag |
| notes           | string | Optional notes |

### Sheet: `Task_History`
| Column          | Type   | Description |
|-----------------|--------|-------------|
| timestamp       | string | When the action was logged |
| task_id         | string | Task id |
| action          | string | worked or completed |
| minutes         | int    | Minutes spent |
| effort          | int    | 1–5 |
| could_be_faster | bool   | Reflection flag |
| notes           | string | Optional notes |

### Analytics Sheets
- `Daily_Summaries`: daily totals, top categories, narrative, suggestions.
- `Weekly_Summaries`: weekly trends, frequent activities, pie data, narrative.
- `Missed_Checkins`: expected vs actual check-ins and largest gaps.
- `Charts_Weekly`: weekly stacked bar source data + chart.
- `Charts_Daily`: daily focus source data + chart.

## Scheduling Model and Failure Modes
The scheduler persists state to `%APPDATA%\HourlyTracker\state.json` and makes decisions via a pure function that is unit-tested.

Key behaviors:
- Normal cadence: prompt every `interval_minutes`.
- Dismissal: auto-snooze for `dismiss_snooze_minutes`.
- Snooze: defers prompts.
- Pause: defers prompts longer.
- Resume/catch-up: after a long gap (sleep/lock/reboot), prompt once with a catch-up dialog instead of spamming.

Failure modes and mitigations:
- Sleep/lock/reboot: detected as a large gap; one catch-up prompt is shown.
- Missed prompts: summarized in `Missed_Checkins`.
- Duplicate prompts: reduced by persisted `last_prompt_at` and a resume guard (`last_resume_at`).
- Excel corruption risk: reduced via lock file and atomic replace.

## Local-Only Privacy Notes
- Default storage is `%USERPROFILE%\Documents\HourlyTracker\time_log.xlsx` (state/locks remain in `%APPDATA%\HourlyTracker`).
- No cloud or telemetry is implemented.
- Optional local LLM mode only uses the `ollama` CLI when enabled.
- A no-network guardrail is enabled by default and blocks most outbound sockets.
- Data remains readable in Excel; consider optional encryption-at-rest if needed.

## Project Structure
- `hourly_tracker/app.py`: tray app entry point.
- `hourly_tracker/scheduler.py`: scheduler and state machine.
- `hourly_tracker/state.py`: persisted scheduler state.
- `hourly_tracker/dialogs.py`: prompt and catch-up dialogs.
- `hourly_tracker/excel_store.py`: Excel persistence, locking, atomic saves.
- `hourly_tracker/analytics.py`: summaries and HTML report.
- `hourly_tracker/excel_formatting.py`: formatting, validation, charts.
- `hourly_tracker/tagging.py`: rule-based tagging + learned overrides.
- `hourly_tracker/llm_ollama.py`: optional Ollama integration.
- `hourly_tracker/no_network.py`: best-effort no-network mode.
- `tests/test_scheduler.py`: unit tests for scheduling logic.
- `electron-app/`: React + Electron UI (separate app).

## Build and Run Plan (Local)

### 1. Create a virtual environment
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 2. Install dependencies
```powershell
pip install -r requirements.txt
```

### 3. Run the tray app
```powershell
python -m hourly_tracker.app
```

### React + Electron UI (separate app)
This is a second UI option that lives in `electron-app/`.

```powershell
cd electron-app
npm install
npm run dev
```

To build a desktop package:
```powershell
npm run build
npm run pack
```

Notes:
- The Electron app runs in the background with a tray icon.
- Use the tray menu to open the main check-in window or the Tasks-only window.
- Electron auto-starts on login and shows the check-in window every 60 minutes (configurable in code).
- Electron uses its own log and report files by default:
  - `%APPDATA%\\HourlyTracker\\time_log_electron.xlsx`
  - `%APPDATA%\\HourlyTracker\\report_electron.html`

The app will:
- Create state in `%APPDATA%\HourlyTracker` (config.json, state.json, locks).
- Create `%USERPROFILE%\Documents\HourlyTracker` with `time_log.xlsx`, `report.html`, and `reflections/`.
- Start prompting on its schedule.

## Configuration
State/config live in `%APPDATA%\HourlyTracker\config.json` (state_dir). User-facing data lives in `%USERPROFILE%\Documents\HourlyTracker` (data_dir) by default.

Common settings:
- `data_dir`: default `%USERPROFILE%\Documents\HourlyTracker` (log, report, reflections).
- `state_dir`: default `%APPDATA%\HourlyTracker` (config, state.json, locks).
- `log_path`: Excel workbook location (defaults to `data_dir/time_log.xlsx`).
- `report_path`: HTML report location (defaults to `data_dir/report.html`).
- `expenses_path`: default `%USERPROFILE%\Desktop\Expenses.xlsx` for the spending prompt.
- `reflection_enabled`: default true.
- `reflection_time_local`: default "23:30" local time.
- `interval_minutes`: default 60.
- `snooze_minutes`: default 10.
- `dismiss_snooze_minutes`: default 10.
- `pause_minutes`: default 60.
- `catch_up_max_hours`: cap for catch-up prompts.
- `llm_enabled`: default false.
- `llm_model`: e.g., `llama3.1:8b`.
- `no_network_mode`: default true.
- `analytics_rules.entry_hours`: default 1.0.
- `analytics_rules.gap_break_hours`: default 2.0.

## Analytics and Reports
Analytics run automatically after entries are logged.

Outputs:
- Excel sheets: daily, weekly, and missed check-ins summaries.
- HTML report: `%USERPROFILE%\Documents\HourlyTracker\report.html` (with `report_updated_at` timestamp).
- Task history tab: included in the HTML report and `Task_History` sheet.

## Optional Local LLM Mode (Ollama)
This mode is strictly local and only runs if enabled and a model is available.

To enable:
1. Install Ollama locally.
2. Pull a model, for example: `ollama pull llama3.1:8b`.
3. Set `llm_enabled` to `true` in `%APPDATA%\HourlyTracker\config.json`.

Example summary prompt used internally:
- "Return strictly valid JSON with keys: narrative (string), suggestions (array of 3 strings)."

Example classifier prompt used internally:
- "Classify the activity into one of the provided categories. Return strictly valid JSON with keys: category, confidence."

If Ollama is missing or the model is unavailable, the app falls back to heuristics.

## Python Version Notes
- The app should run on Python 3.12+.
- **Packaging with PyInstaller requires Python 3.13 or 3.12.** PyInstaller does not currently support Python 3.14, so use 3.13/3.12 for building the `.exe`.
- `win10toast` is optional. If you want toast notifications, install it manually:
  `pip install win10toast`

## Packaging Into a Windows App (PyInstaller)

### Build command
```powershell
pyinstaller --noconfirm --clean --name HourlyTracker --windowed --onefile run_hourly_tracker.py
```

Notes:
- The generated executable will be in `dist\HourlyTracker.exe`.
- Config/state stay in `%APPDATA%\HourlyTracker`; log/report live in `%USERPROFILE%\Documents\HourlyTracker`.

### Installer-friendly layout (suggested)
- `dist\HourlyTracker.exe`
- `dist\README.txt`
- `dist\LICENSE.txt`

## Auto-Start on Login (Optional)
Use the helper script to add a Startup shortcut:

```powershell
powershell -ExecutionPolicy Bypass -File packaging/enable_autostart.ps1 -ExePath "C:\path\to\HourlyTracker.exe"
```

If you built with the default script, the exe is usually:
`dist\HourlyTracker.exe`

## UX Features Included
- Always-on-top prompt dialog.
- Enter submits, Escape dismisses.
- Dismiss auto-snoozes.
- "I'm in class/driving" button logs a default activity.
- Category suggestion updates while typing, with a confidence hint.
- Tray toasts when a prompt is due.

## Security and Privacy Checklist
Threat model considerations:
- Local attacker with access to your user profile.
- Malware running as your user.
- Shared PC with other accounts.

Checklist:
- Data location: verify `%USERPROFILE%\\Documents\\HourlyTracker` (data) and `%APPDATA%\\HourlyTracker` (state) are acceptable.
- Backups: regularly back up `time_log.xlsx`.
- Retention: consider periodic archiving or pruning.
- Encryption at rest: consider wrapping the workbook in encrypted storage (for example, BitLocker or an encrypted container).
- File permissions: ensure your Windows account is protected by a password and disk encryption where possible.
- No telemetry: confirmed by code inspection; no analytics services are used.
- No-network mode: enabled by default via `no_network.py` guardrails.

## Running Tests
```powershell
python -m unittest discover -s tests -p "test_*.py"
```

## Notes and Limitations
- The no-network guardrail is best-effort and not a hard security boundary.
- Tkinter dialogs run on a dedicated thread and are intentionally simple.
- Excel chart and validation behavior can vary slightly across Excel versions.
