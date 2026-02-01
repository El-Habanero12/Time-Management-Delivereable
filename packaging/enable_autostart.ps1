param(
  [string]$ExePath = ""
)

$ErrorActionPreference = "Stop"

$startup = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
$shortcut = Join-Path $startup "HourlyTracker.lnk"

$ws = New-Object -ComObject WScript.Shell

if (-not $ExePath) {
  try {
    $ExePath = (Resolve-Path ".\\dist\\HourlyTracker.exe").Path
  } catch {
    throw "ExePath not provided and default .\\dist\\HourlyTracker.exe not found. Pass -ExePath to the script."
  }
}

if (-not (Test-Path $ExePath)) {
  throw "ExePath not found: $ExePath"
}

$s = $ws.CreateShortcut($shortcut)
$s.TargetPath = $ExePath
$s.WorkingDirectory = Split-Path -Path $ExePath
$s.Save()

Write-Host "Created startup shortcut: $shortcut"
