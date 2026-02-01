# HourlyTracker

## For users
- Download the packaged `HourlyTracker.exe` (PyInstaller build).
- Run the EXE; no Python required.
- Data locations (NORMAL mode):
  - State/config/logs: `%APPDATA%\HourlyTracker`
  - User files: `%USERPROFILE%\Documents\HourlyTracker`
    - `time_log.xlsx` (first run copies from template)
    - `Expenses.xlsx` (Tracker sheet; first run copies from template)
    - `reflections\*.docx`
- TEST mode: set environment variable `HOURLYTRACKER_PROFILE=TEST` before launching; folders become `%APPDATA%\HourlyTracker_TEST` and `Documents\HourlyTracker_TEST`.

## For developers
- Create a virtual environment and install requirements:
  - `py -3 -m venv .venv`
  - `.venv\Scripts\activate`
  - `pip install -r requirements.txt`
- Run the tray app in dev: `python run_hourly_tracker.py`
- Build a standalone EXE:
  - `pyinstaller HourlyTracker.spec`
  - Artifacts appear under `dist/HourlyTracker/` (onedir) or `dist/HourlyTracker.exe` (onefile, if configured).
