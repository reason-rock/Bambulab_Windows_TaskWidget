$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonExe = "$env:LOCALAPPDATA\Python\pythoncore-3.14-64\python.exe"
$projectIcon = Join-Path $scriptDir "app_icon.ico"
$distDir = Join-Path $scriptDir "dist_release"
$workDir = Join-Path $scriptDir "build_release"

if (-not (Test-Path $pythonExe)) {
    throw "python.exe를 찾지 못했습니다: $pythonExe"
}

$args = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--windowed",
    "--distpath", $distDir,
    "--workpath", $workDir,
    "--name", "Bambu Monitor",
    "--collect-all", "winotify",
    "--hidden-import", "pystray._win32",
    "--hidden-import", "PIL._tkinter_finder",
    "bbmonitor.py"
)

if (Test-Path $projectIcon) {
    $args += @("--icon", $projectIcon)
    Write-Host "아이콘 사용: $projectIcon"
} else {
    Write-Host "경고: app_icon.ico가 없습니다. 기본 아이콘으로 빌드합니다."
}

Push-Location $scriptDir
try {
    & $pythonExe @args
} finally {
    Pop-Location
}
