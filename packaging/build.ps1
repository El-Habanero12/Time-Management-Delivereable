param(
  [string]$Python = "python"
)

$ErrorActionPreference = "Stop"

function Get-PythonVersion([string]$Py) {
  $ver = & $Py -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}')"
  return $ver
}

$version = Get-PythonVersion $Python
$parts = $version -split '\.'
$major = [int]$parts[0]
$minor = [int]$parts[1]

if ($major -ge 3 -and $minor -ge 14) {
  throw "PyInstaller does not yet support Python $version. Install Python 3.13 or 3.12 for packaging."
}

& $Python -m pip install -r requirements.txt
& $Python -m PyInstaller --noconfirm --name HourlyTracker --windowed --onefile hourly_tracker\app.py
if ($LASTEXITCODE -ne 0) {
  throw "PyInstaller failed with exit code $LASTEXITCODE"
}

$exe = Join-Path (Resolve-Path ".\\dist").Path "HourlyTracker.exe"
if (-not (Test-Path $exe)) {
  throw "Build finished but exe not found at $exe"
}
Write-Host "Built $exe"
