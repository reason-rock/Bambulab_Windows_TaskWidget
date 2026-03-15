param(
    [switch]$StartupOnly
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$targetScript = Join-Path $scriptDir "bbmonitor.py"

if (-not (Test-Path $targetScript)) {
    throw "bbmonitor.py를 찾을 수 없습니다: $targetScript"
}

function Resolve-Pythonw {
    $candidates = @(
        "$env:LOCALAPPDATA\Python\pythoncore-3.14-64\pythonw.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python314\pythonw.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python313\pythonw.exe"
    )

    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCmd) {
        $candidate = (Join-Path (Split-Path $pythonCmd.Source -Parent) "pythonw.exe")
        if ($candidate -and ($candidate -notmatch "WindowsApps")) {
            $candidates += $candidate
        }
    }

    $pyCmd = Get-Command py -ErrorAction SilentlyContinue
    if ($pyCmd) {
        $candidates += "$env:LOCALAPPDATA\Programs\Python\Python314\pythonw.exe"
        $candidates += "$env:LOCALAPPDATA\Programs\Python\Python313\pythonw.exe"
        $candidates += "$env:LOCALAPPDATA\Python\pythoncore-3.14-64\pythonw.exe"
    }

    $candidates = $candidates | Where-Object { $_ -and ($_ -notmatch "WindowsApps") } | Select-Object -Unique

    foreach ($path in $candidates) {
        if (Test-Path $path) {
            return $path
        }
    }

    throw "pythonw.exe를 찾지 못했습니다. Python 설치 경로를 확인하세요."
}

function Resolve-DesktopPath {
    $paths = @(
        [Environment]::GetFolderPath("Desktop"),
        "$env:USERPROFILE\OneDrive - Tommoro Robotics\바탕 화면",
        "$env:USERPROFILE\Desktop"
    ) | Where-Object { $_ }

    foreach ($p in $paths | Select-Object -Unique) {
        if (Test-Path $p) {
            return $p
        }
    }

    throw "Desktop 경로를 찾지 못했습니다."
}

function New-AppShortcut {
    param(
        [Parameter(Mandatory=$true)][string]$ShortcutPath,
        [Parameter(Mandatory=$true)][string]$PythonwPath,
        [Parameter(Mandatory=$true)][string]$ScriptPath,
        [Parameter(Mandatory=$true)][string]$WorkingDir
    )

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $PythonwPath
    $shortcut.Arguments = "`"$ScriptPath`""
    $shortcut.WorkingDirectory = $WorkingDir
    $shortcut.IconLocation = "$env:SystemRoot\System32\imageres.dll,95"
    $shortcut.WindowStyle = 7
    $shortcut.Description = "Bambu Monitor"
    $shortcut.Save()
}

$pythonwPath = Resolve-Pythonw
$desktopPath = Resolve-DesktopPath
$startupPath = [Environment]::GetFolderPath("Startup")

$desktopShortcut = Join-Path $desktopPath "Bambu Monitor.lnk"
$startupShortcut = Join-Path $startupPath "Bambu Monitor.lnk"

if (-not $StartupOnly) {
    New-AppShortcut -ShortcutPath $desktopShortcut -PythonwPath $pythonwPath -ScriptPath $targetScript -WorkingDir $scriptDir
    Write-Host "Desktop shortcut created: $desktopShortcut"
}

New-AppShortcut -ShortcutPath $startupShortcut -PythonwPath $pythonwPath -ScriptPath $targetScript -WorkingDir $scriptDir
Write-Host "Startup shortcut created: $startupShortcut"
Write-Host "Pythonw target: $pythonwPath"
Write-Host "Done."
