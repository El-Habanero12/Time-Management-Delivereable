$ErrorActionPreference = "Stop"

function Resolve-Python {
    $candidates = @("py -3.13", "py -3", "python")
    foreach ($cmd in $candidates) {
        try {
            & $cmd --version *> $null
            if ($LASTEXITCODE -eq 0) {
                return $cmd
            }
        } catch {
            continue
        }
    }
    throw "Python launcher not found. Please install Python 3.10+."
}

$python = Resolve-Python
Write-Host "Using Python: $python"

# Ensure venv
if (-not (Test-Path ".venv/Scripts/Activate.ps1")) {
    Write-Host "Creating virtual environment..."
    & $python -m venv .venv
}

& ".\.venv\Scripts\python.exe" -m pip install --upgrade pip
& ".\.venv\Scripts\python.exe" -m pip install -r requirements.txt

Write-Host "Running PyInstaller..."
& ".\.venv\Scripts\python.exe" -m PyInstaller --clean --noconfirm HourlyTracker.spec

Write-Host "Build complete. Output:"
Write-Host "  dist\\HourlyTracker\\HourlyTracker.exe (onedir)"
Write-Host "  or dist\\HourlyTracker.exe (onefile if configured in spec)"
