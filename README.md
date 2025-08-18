# AutoHotkey v2 Window Cycling

A fast and reliable AutoHotkey v2 script for launching and cycling through application windows using keyboard shortcuts.

## Features

- **Smart window cycling** - Skip currently active windows and jump to the next one
- **Non-blocking launch focus** - Applications are launched and focused without blocking other hotkeys
- **Robust window detection** - Handles minimized, cloaked, and closed windows gracefully
- **Cached window snapshots** - 5-second caching reduces system calls for better performance
- **Focus stealing prevention** - Cancels launch timers when switching to other applications

## Default Hotkeys

- **Alt + 1** - Windows Terminal
- **Alt + 2** - Visual Studio
- **Alt + 3** - Vivaldi Browser
- **Alt + 4** - ChatGPT Desktop
- **Alt + 7** - Obsidian
- **Alt + 8** - ProtonMail
- **Win + 1** - File Explorer
- **Win + 2** - Microsoft Edge
- **Win + 3** - Microsoft Outlook
- **Win + 4** - Microsoft Teams
- **Win + 0** - Task Manager

## How It Works

1. **Single press** - Activates the application or launches it if not running
2. **Multiple windows** - Cycles through all windows of the application
3. **Smart cycling** - If the target app is already active, immediately jumps to the next window
4. **Dead handle cleanup** - Automatically removes closed windows from the cycling list

## Requirements

- AutoHotkey v2.0 or later
- Windows 10/11 (uses DWM cloaking detection)

## Customization

Edit the hotkey definitions at the bottom of the script to match your preferred applications and hotkeys:

```autohotkey
YourHotkey::LaunchOrCycle("your-process.exe", "launch-command", "appKey")
```

Parameters:
- `processName` - The executable name to detect windows
- `launchCommand` - Command to launch the application  
- `appKey` - Unique identifier for the application

The hotkey can be any combination supported by AutoHotkey (e.g., `!1`, `#1`, `^j`, `F1`, etc.).

## Installation

1. Install [AutoHotkey v2](https://www.autohotkey.com/v2/)
2. Download or clone this repository
3. Run `Open_And_Cycle_Apps.ahk`
4. Customize the hotkeys and applications as needed

## Technical Details

- Uses traditional Alt modifier syntax (!n::) for better application hotkey priority
- Implements timer-based window polling for reliable launch focus
- Filters out DWM-cloaked windows using Windows API calls
- Maintains per-application state for efficient window management
- Handles Windows foreground lock restrictions gracefully

## License

This project is open source and available under the MIT License.