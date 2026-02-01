$ErrorActionPreference = "Stop"

$paths = @(
    "build",
    "dist",
    "build_tmp",
    "dist_tmp",
    "*.tmp",
    "time_log.xlsx",
    "Expenses.xlsx",
    "report.html",
    "reflections"
)

foreach ($p in $paths) {
    Get-ChildItem -Path $p -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "Clean completed."
