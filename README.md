# Bambu Monitor

A small Windows tray app for monitoring Bambu Lab printers over MQTT.

It shows printer progress as a battery-style system tray icon, changes color by state, and displays Windows notifications when a print starts or finishes.

## Disclaimer

This is an unofficial community project.

- It is not affiliated with, endorsed by, or supported by Bambu Lab.
- It may rely on behaviors, endpoints, or protocols that can change without notice.
- Use it at your own risk.
- You are responsible for complying with Bambu Lab terms, policies, and local laws in your region.
- Do not use this repository to publish or share real credentials, tokens, or device identifiers.
- This repository should be treated as Windows-only unless you port and validate the tray and notification behavior yourself.

## Features

- Live printer state updates over Bambu Lab MQTT
- Multi-printer support with tray icon rotation every 5 seconds
- Battery-style tray icon with progress number
- State-based colors for printing, idle, paused, finished, failed, and disconnected states
- Windows popup notifications for print start, print finish, and print failure
- Progress milestone notifications (every 25%)
- Automatic MQTT reconnection with exponential backoff (up to 20 attempts)
- Connection status display in tray tooltip (reconnecting attempt count)
- Settings GUI (tkinter) for credential and printer management
- Optional Windows taskbar progress bar (PyTaskbar)
- Single-instance protection
- Optional desktop shortcut and auto-start registration
- Standalone `.exe` build via PyInstaller
- Separate popup test script for notification debugging

## Current Behavior

Tray icon colors:

- `printing`: red
- `idle` / `finished`: green
- `paused`: amber
- `failed`: dark red
- `disconnected`: gray

Notifications:

- Print start: long popup
- Print finish: long popup
- Print failure: long popup
- Progress milestones (25%, 50%, 75%): short popup
- MQTT reconnection failed (20 attempts): long popup

MQTT reconnection:

- On disconnect, automatically retries with exponential backoff (2s -> 4s -> 8s -> ... -> max 120s)
- Printer state is preserved across reconnections
- Tray tooltip shows `[재연결 N/20]` during reconnect, `[연결 끊김]` after max attempts

## Requirements

- Windows
- Python 3.11 or newer recommended
- A Bambu Lab account with a linked printer
- MQTT credentials for the target printer

## Install

```powershell
git clone <your-repo-url>
cd Bambulab_Windows_TaskWidget
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

`PyTaskbar` is optional. The main tray icon works without it.

## Configuration

The app supports two configuration methods, loaded in this order:

1. **Environment variables** (`.env` file) — used as defaults
2. **Config file** (`bambu_monitor_config.json`) — takes precedence, created by the settings GUI

On first run, if no valid config is found, the settings dialog opens automatically.

### Environment Variables

Copy `.env.example` to `.env`:

```powershell
Copy-Item .env.example .env
```

Then fill in your values:

```env
BBL_USER_ID=your_user_id
BBL_ACCESS_TOKEN=your_access_token
BBL_DEVICE_ID=your_device_id
BBL_EMAIL=your_bambu_email
BBL_PASSWORD=your_bambu_password
```

Notes:

- `BBL_USER_ID`, `BBL_ACCESS_TOKEN`, and `BBL_DEVICE_ID` are used by `bbmonitor.py` and `mqttmonitor.py`.
- `BBL_EMAIL` and `BBL_PASSWORD` are used by helper login scripts such as `bblab.py` and `bblogin.py`.
- Do not commit `.env`.

### Settings GUI

You can also configure credentials and printers through the built-in settings dialog:

- Opens automatically on first run
- Accessible from the tray context menu ("설정 열기")
- Supports multiple printers in `별명 | DEVICE_ID` format
- Settings are saved to `bambu_monitor_config.json`

## Run

Run the monitor directly:

```powershell
python .\bbmonitor.py
```

If the required tray dependencies are installed, the app will appear in the Windows system tray.

## Build Standalone EXE

Package the app into a standalone `.exe` using PyInstaller:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
```

The output is located in `dist_release/Bambu Monitor/`.

## Install Desktop Shortcut and Auto Start

Create both a desktop shortcut and a Startup shortcut:

```powershell
powershell -ExecutionPolicy Bypass -File .\install_shortcuts.ps1
```

Remove those shortcuts:

```powershell
powershell -ExecutionPolicy Bypass -File .\uninstall_shortcuts.ps1
```

## Test Notifications Only

To test Windows popup notifications without connecting to MQTT:

```powershell
python .\test_popup.py
```

This sends two test notifications:

- popup test for print start
- popup test for print completion

## Project Files

- `bbmonitor.py`: main Windows tray monitor
- `mqttmonitor.py`: simpler MQTT monitor and debug script
- `bblab.py`: login helper
- `bblogin.py`: login and device listing helper
- `test_popup.py`: popup-only test script
- `build_exe.ps1`: package app into standalone `.exe` via PyInstaller
- `install_shortcuts.ps1`: create desktop and startup shortcuts
- `uninstall_shortcuts.ps1`: remove created shortcuts

## Troubleshooting

If the tray icon does not appear:

- check hidden tray icons (`^`)
- confirm `pystray` and `pillow` are installed
- check Windows notification area settings

If notifications do not appear:

- check Windows notification settings
- disable Focus Assist / Do Not Disturb
- run `python .\test_popup.py`

If the app appears to start more than once:

- the app includes a single-instance lock
- exit old instances from the tray menu before re-testing

If MQTT keeps disconnecting:

- check your network connection
- verify your credentials are still valid (re-run `bblogin.py` to refresh)
- the app will automatically attempt to reconnect up to 20 times
- after 20 failed attempts, a notification will prompt you to restart

## Security

This repository should not contain live credentials.

Before publishing:

- keep `.env` out of version control
- do not commit `bambu_monitor_config.json` (contains tokens)
- rotate any credentials that were previously hardcoded or shared
- review commit history if secrets were ever committed earlier

## License

This project is released under the MIT License. See `LICENSE`.
