#Requires AutoHotkey v2.0

; AutoHotkey v2 Window Cycling Script
; Provides keyboard shortcuts for launching and cycling through application windows
; with snapshot-based caching and smart focus management

; Global timer storage for launch operations
g_activeTimers := Map()

; Track which app was last launched per appKey (for multi-app keybinds)
g_lastLaunched := Map()

; Check if window is cloaked by DWM (hidden by Windows)
IsCloaked(hwnd) {
    static DWMWA_CLOAKED := 14
    flag := 0
    ok := DllCall("dwmapi\DwmGetWindowAttribute"
        , "ptr", hwnd, "int", DWMWA_CLOAKED
        , "int*", &flag, "int", 4, "int")
    return (ok = 0) && (flag != 0)
}

; Validate if window exists and is interactable
IsRealWindow(hwnd) {
    try {
        if !WinExist("ahk_id " hwnd)
            return false
        if IsCloaked(hwnd)
            return false
        return true
    }
    catch {
        return false
    }
}

; Cancel pending launch timers to prevent focus stealing
CancelAllLaunchTimers() {
    global g_activeTimers
    for appKey, timer in g_activeTimers.Clone() {
        SetTimer(timer, 0)
        g_activeTimers.Delete(appKey)
    }
}

; Main function to launch or cycle through application windows
; Supports multiple apps per keybind via arrays (e.g., ["devenv.exe", "Code.exe"])
LaunchOrCycle(processNames, launchCommands, appKey) {
    global g_lastLaunched

    ; Per-app state storage
    static snapshots := Map()

    ; Cancel pending timers to prevent focus stealing
    CancelAllLaunchTimers()

    ; Initialize app state if needed
    if (!snapshots.Has(appKey)) {
        snapshots[appKey] := {
            windowSnapshot: [],
            lastActivated: 0,
            snapshotTime: 0
        }
    }

    appData := snapshots[appKey]

    ; Refresh window list if expired (5 second cache)
    currentTime := A_TickCount
    if (currentTime - appData.snapshotTime > 5000 || appData.windowSnapshot.Length = 0) {
        appData.windowSnapshot := []

        ; Build fresh window list from all process names
        for procName in processNames {
            if (procName = "explorer.exe") {
                wins := WinGetList("ahk_class CabinetWClass")
            } else {
                wins := WinGetList("ahk_exe " . procName)
            }

            ; Filter for valid windows
            for id in wins {
                if (IsRealWindow(id)) {
                    appData.windowSnapshot.Push(id)
                }
            }
        }

        appData.snapshotTime := currentTime
        appData.lastActivated := 0
    }

    ; Launch if no windows exist (use last launched app, or first)
    if (appData.windowSnapshot.Length = 0) {
        launchIndex := g_lastLaunched.Has(appKey) ? g_lastLaunched[appKey] : 1
        LaunchAndFocus(launchCommands[launchIndex], processNames[launchIndex], appKey, launchIndex)
        return
    }
    
    ; Single window - just activate
    if (appData.windowSnapshot.Length = 1) {
        ActivateWindow(appData.windowSnapshot[1])
        return
    }
    
    ; Smart cycling - skip currently active window
    activeWindow := WinGetID("A")
    currentIndex := 0
    
    ; Check if active window belongs to this app
    for i, hwnd in appData.windowSnapshot {
        if (hwnd = activeWindow) {
            currentIndex := i
            break
        }
    }
    
    ; Jump to next if current app is active, otherwise cycle normally
    if (currentIndex > 0) {
        appData.lastActivated := (currentIndex >= appData.windowSnapshot.Length) ? 1 : currentIndex + 1
    } else {
        appData.lastActivated := (appData.lastActivated >= appData.windowSnapshot.Length) ? 1 : appData.lastActivated + 1
    }
    
    ; Activate target and handle closed windows
    target := appData.windowSnapshot[appData.lastActivated]
    ActivateWindow(target)
    
    ; Clean up closed windows and retry
    if (WinGetID("A") != target) {
        appData.windowSnapshot.RemoveAt(appData.lastActivated)
        appData.snapshotTime := A_TickCount  ; Mark as fresh to prevent immediate rebuild
        if (appData.windowSnapshot.Length > 0) {
            ; Adjust index after removal
            if (appData.lastActivated > appData.windowSnapshot.Length) {
                appData.lastActivated := 1
            }
            ActivateWindow(appData.windowSnapshot[appData.lastActivated])
        }
    }
}

; Activate window with proper restoration and focus handling
ActivateWindow(hwnd) {
    if !WinExist("ahk_id " hwnd)
        return
    
    ; Restore minimized windows
    if (WinGetMinMax("ahk_id " hwnd) = -1) {
        WinRestore("ahk_id " hwnd)
        Sleep(20)
    }
    
    ; Handle Windows focus restrictions
    try DllCall("user32\AllowSetForegroundWindow", "uint", -1)
    WinActivate("ahk_id " hwnd)
}

; Launch application with non-blocking focus polling
LaunchAndFocus(launchCommand, processName, appKey, launchIndex := 1) {
    global g_activeTimers, g_lastLaunched

    ; Clean up existing timer
    if (g_activeTimers.Has(appKey)) {
        SetTimer(g_activeTimers[appKey], 0)
        g_activeTimers.Delete(appKey)
    }

    ; Start the application
    pid := Run(launchCommand)
    launchTime := A_TickCount

    ; Poll for window appearance
    PollForWindow() {

        ; Try exact PID match first
        for hwnd in WinGetList("ahk_pid " pid) {
            if (IsRealWindow(hwnd)) {
                ActivateWindow(hwnd)
                g_lastLaunched[appKey] := launchIndex
                SetTimer(PollForWindow, 0)
                g_activeTimers.Delete(appKey)
                return
            }
        }

        ; Fallback to process name after 400ms
        if (A_TickCount - launchTime > 400) {
            for hwnd in WinGetList("ahk_exe " processName) {
                if (IsRealWindow(hwnd)) {
                    ActivateWindow(hwnd)
                    g_lastLaunched[appKey] := launchIndex
                    SetTimer(PollForWindow, 0)
                    g_activeTimers.Delete(appKey)
                    return
                }
            }
        }

        ; Stop polling after 5 seconds
        if (A_TickCount - launchTime > 5000) {
            SetTimer(PollForWindow, 0)
            g_activeTimers.Delete(appKey)
        }
    }

    ; Start polling timer
    g_activeTimers[appKey] := PollForWindow
    SetTimer(PollForWindow, 75)
}

; Launch new window: opens the other app if only one is running, or new window of focused app
LaunchNew(processNames, launchCommands, appKey) {
    ; Count which apps have windows
    appWindowCounts := []
    for procName in processNames {
        count := 0
        if (procName = "explorer.exe") {
            wins := WinGetList("ahk_class CabinetWClass")
        } else {
            wins := WinGetList("ahk_exe " . procName)
        }
        for id in wins {
            if (IsRealWindow(id))
                count++
        }
        appWindowCounts.Push(count)
    }

    ; For dual-app: if only one app running, launch the other
    if (processNames.Length = 2) {
        if (appWindowCounts[1] > 0 && appWindowCounts[2] = 0) {
            Run(launchCommands[2])
            return
        }
        if (appWindowCounts[2] > 0 && appWindowCounts[1] = 0) {
            Run(launchCommands[1])
            return
        }
    }

    ; Both running (or single-app): launch new window of focused app
    activeWindow := WinGetID("A")
    try activeExe := WinGetProcessName("ahk_id " activeWindow)
    catch
        activeExe := ""

    ; Find which app is focused and launch that
    for i, procName in processNames {
        if (procName = activeExe) {
            Run(launchCommands[i])
            return
        }
    }

    ; Fallback: launch first app
    Run(launchCommands[1])
}

; Alt + 1: Windows Terminal
!1::LaunchOrCycle(["WindowsTerminal.exe"], ["wt.exe"], "terminal")

; Alt + 2: Visual Studio + VSCode
!2::LaunchOrCycle(["devenv.exe", "Code.exe"], ["devenv.exe", "C:\Users\" . A_UserName . "\AppData\Local\Programs\Microsoft VS Code\Code.exe"], "ide")

; Alt + 3: Vivaldi
!3::LaunchOrCycle(["vivaldi.exe"], ["vivaldi.exe"], "vivaldi")

; Alt + 4: ChatGPT + Claude Desktop
!4::LaunchOrCycle(["ChatGPT.exe", "Claude.exe"], ["ChatGPT.exe", "C:\Users\" . A_UserName . "\AppData\Local\AnthropicClaude\claude.exe"], "ai")

; Alt + 5: GitKraken
!5::LaunchOrCycle(["GitKraken.exe"], ["C:\Users\" . A_UserName . "\AppData\Local\gitkraken\gitkraken.exe"], "gitkraken")

; Alt + 6: Insomnia
!6::LaunchOrCycle(["Insomnia.exe"], ["C:\Users\" . A_UserName . "\AppData\Local\insomnia\Insomnia.exe"], "insomnia")

; Alt + 7: Obsidian
!7::LaunchOrCycle(["Obsidian.exe"], ["obsidian://"], "obsidian")

; Alt + 8: ProtonMail
!8::LaunchOrCycle(["Proton Mail.exe"], ["C:\Users\DagOleBergersenHimle\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Proton\Proton Mail.lnk"], "protonmail")

; Win + 1: File Explorer
#1::LaunchOrCycle(["explorer.exe"], ["explorer.exe"], "explorer")

; Win + 2: Microsoft Edge
#2::LaunchOrCycle(["msedge.exe"], ["msedge.exe"], "edge")

; Win + 3: Microsoft Outlook (New Outlook)
#3::LaunchOrCycle(["olk.exe"], ["olk.exe"], "outlook")

; Win + 4: Microsoft Teams
#4::LaunchOrCycle(["ms-teams.exe"], ["ms-teams.exe"], "teams")

; Win + 0: Task Manager
#0::LaunchOrCycle(["Taskmgr.exe"], ["taskmgr.exe"], "taskmgr")

; ============================================
; Shift+Alt: Open new window / launch other app
; ============================================

; Shift+Alt + 1: New Terminal window
+!1::LaunchNew(["WindowsTerminal.exe"], ["wt.exe"], "terminal")

; Shift+Alt + 2: New VS/VSCode (or launch the other)
+!2::LaunchNew(["devenv.exe", "Code.exe"], ["devenv.exe", "C:\Users\" . A_UserName . "\AppData\Local\Programs\Microsoft VS Code\Code.exe"], "ide")

; Shift+Alt + 3: New Vivaldi window
+!3::LaunchNew(["vivaldi.exe"], ["vivaldi.exe"], "vivaldi")

; Shift+Alt + 4: New ChatGPT/Claude (or launch the other)
+!4::LaunchNew(["ChatGPT.exe", "Claude.exe"], ["ChatGPT.exe", "C:\Users\" . A_UserName . "\AppData\Local\AnthropicClaude\claude.exe"], "ai")

; Shift+Alt + 5: New GitKraken window
+!5::LaunchNew(["GitKraken.exe"], ["C:\Users\" . A_UserName . "\AppData\Local\gitkraken\gitkraken.exe"], "gitkraken")

; Shift+Alt + 6: New Insomnia window
+!6::LaunchNew(["Insomnia.exe"], ["C:\Users\" . A_UserName . "\AppData\Local\insomnia\Insomnia.exe"], "insomnia")

; Shift+Alt + 7: New Obsidian window
+!7::LaunchNew(["Obsidian.exe"], ["obsidian://"], "obsidian")

; Shift+Alt + 8: New ProtonMail window
+!8::LaunchNew(["Proton Mail.exe"], ["C:\Users\DagOleBergersenHimle\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Proton\Proton Mail.lnk"], "protonmail")

; Shift+Win + 1: New Explorer window
+#1::LaunchNew(["explorer.exe"], ["explorer.exe"], "explorer")

; Shift+Win + 2: New Edge window
+#2::LaunchNew(["msedge.exe"], ["msedge.exe"], "edge")

; Shift+Win + 3: New Outlook window
+#3::LaunchNew(["olk.exe"], ["olk.exe"], "outlook")

; Shift+Win + 4: New Teams window
+#4::LaunchNew(["ms-teams.exe"], ["ms-teams.exe"], "teams")

; Shift+Win + 0: New Task Manager window
+#0::LaunchNew(["Taskmgr.exe"], ["taskmgr.exe"], "taskmgr")
