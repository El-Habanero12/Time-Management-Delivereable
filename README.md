# HourlyTracker

## For Users
- Download and run `HourlyTracker.exe` (no Python required).
- Storage (NORMAL mode):
  - State/config/logs: `%APPDATA%\HourlyTracker`
  - User files: `%USERPROFILE%\Documents\HourlyTracker`
    - `time_log.xlsx` (copied from template on first run)
    - `Expenses.xlsx` (Tracker sheet; copied from template on first run)
    - `reflections\*.docx`
- TEST mode: set `HOURLYTRACKER_PROFILE=TEST` before launching; folders become `%APPDATA%\HourlyTracker_TEST` and `Documents\HourlyTracker_TEST`.
- If Excel says a workbook is open/locked, close it and retry; the tray app will stay running and notify you.

## For Developers
- Create a venv and install deps:
  - `py -3 -m venv .venv`
  - `.venv\Scripts\activate`
  - `pip install -r requirements.txt`
- Run in dev: `python run_hourly_tracker.py`
- Build: `pyinstaller HourlyTracker.spec` or `powershell -ExecutionPolicy Bypass -File scripts/build.ps1`
- Run built EXE in TEST mode by setting `HOURLYTRACKER_PROFILE=TEST` in your shell before launching.
