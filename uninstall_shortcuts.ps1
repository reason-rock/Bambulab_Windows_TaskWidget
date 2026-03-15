$ErrorActionPreference = "Stop"

$desktopPath = [Environment]::GetFolderPath("Desktop")
$startupPath = [Environment]::GetFolderPath("Startup")

$desktopShortcut = Join-Path $desktopPath "Bambu Monitor.lnk"
$startupShortcut = Join-Path $startupPath "Bambu Monitor.lnk"

foreach ($path in @($desktopShortcut, $startupShortcut)) {
    if (Test-Path $path) {
        Remove-Item -Path $path -Force
        Write-Host "Removed: $path"
    } else {
        Write-Host "Not found: $path"
    }
}

Write-Host "Done."
