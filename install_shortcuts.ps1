param(
    [switch]$StartupOnly
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$targetScript = Join-Path $scriptDir "bbmonitor.py"
$releaseExe = Join-Path $scriptDir "dist_release\Bambu Monitor\Bambu Monitor.exe"
$targetExe = Join-Path $scriptDir "dist\Bambu Monitor\Bambu Monitor.exe"
$appUserModelId = "reason_rock.BambuMonitor"

if ((-not (Test-Path $targetScript)) -and (-not (Test-Path $targetExe)) -and (-not (Test-Path $releaseExe))) {
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
        [Parameter(Mandatory=$true)][string]$TargetPath,
        [string]$Arguments,
        [Parameter(Mandatory=$true)][string]$WorkingDir
    )

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.Arguments = $Arguments
    $shortcut.WorkingDirectory = $WorkingDir
    $shortcut.IconLocation = $TargetPath
    $shortcut.WindowStyle = 7
    $shortcut.Description = "Bambu Monitor"
    try {
        $shortcut | Add-Member -NotePropertyName AppUserModelID -NotePropertyValue $appUserModelId -Force
    } catch {
    }
    $shortcut.Save()
}

$desktopPath = Resolve-DesktopPath
$startupPath = [Environment]::GetFolderPath("Startup")

$desktopShortcut = Join-Path $desktopPath "Bambu Monitor.lnk"
$startupShortcut = Join-Path $startupPath "Bambu Monitor.lnk"

if (Test-Path $releaseExe) {
    $shortcutTarget = $releaseExe
    $shortcutArguments = ""
    Write-Host "Using packaged executable: $releaseExe"
} elseif (Test-Path $targetExe) {
    $shortcutTarget = $targetExe
    $shortcutArguments = ""
    Write-Host "Using packaged executable: $targetExe"
} else {
    $pythonwPath = Resolve-Pythonw
    $shortcutTarget = $pythonwPath
    $shortcutArguments = "`"$targetScript`""
    Write-Host "Using pythonw target: $pythonwPath"
}

if (-not $StartupOnly) {
    New-AppShortcut -ShortcutPath $desktopShortcut -TargetPath $shortcutTarget -Arguments $shortcutArguments -WorkingDir $scriptDir
    Write-Host "Desktop shortcut created: $desktopShortcut"
}

New-AppShortcut -ShortcutPath $startupShortcut -TargetPath $shortcutTarget -Arguments $shortcutArguments -WorkingDir $scriptDir
Write-Host "Startup shortcut created: $startupShortcut"
Write-Host "Done."
