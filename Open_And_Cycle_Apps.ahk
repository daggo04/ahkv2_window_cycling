#Requires AutoHotkey v2.0

; AutoHotkey v2 Window Cycling Script
; Provides keyboard shortcuts for launching and cycling through application windows
; with snapshot-based caching and smart focus management

; Global timer storage for launch operations
g_activeTimers := Map()

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
LaunchOrCycle(processName, launchCommand, appKey) {
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
        ; Build fresh window list
        if (processName = "explorer.exe") {
            wins := WinGetList("ahk_class CabinetWClass")
        } else {
            wins := WinGetList("ahk_exe " . processName)
        }
        appData.windowSnapshot := []
        
        ; Filter for valid windows
        for id in wins {
            if (IsRealWindow(id)) {
                appData.windowSnapshot.Push(id)
            }
        }
        
        appData.snapshotTime := currentTime
        appData.lastActivated := 0
    }
    
    ; Launch if no windows exist
    if (appData.windowSnapshot.Length = 0) {
        LaunchAndFocus(launchCommand, processName, appKey)
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
LaunchAndFocus(launchCommand, processName, appKey) {
    global g_activeTimers
    
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

; Alt + 1: Windows Terminal
!1::LaunchOrCycle("WindowsTerminal.exe", "wt.exe", "terminal")

; Alt + 2: Visual Studio
!2::LaunchOrCycle("devenv.exe", "devenv.exe", "visualstudio")

; Alt + 3: Vivaldi
!3::LaunchOrCycle("vivaldi.exe", "vivaldi.exe", "vivaldi")

; Alt + 4: ChatGPT Desktop
!4::LaunchOrCycle("ChatGPT.exe", "ChatGPT.exe", "chatgpt")

; Alt + 7: Obsidian
!7::LaunchOrCycle("Obsidian.exe", "obsidian://", "obsidian")

; Alt + 8: ProtonMail
!8::LaunchOrCycle("Proton Mail.exe", "C:\Users\DagOleBergersenHimle\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Proton\Proton Mail.lnk", "protonmail")

; Win + 1: File Explorer
#1::LaunchOrCycle("explorer.exe", "explorer.exe", "explorer")

; Win + 2: Microsoft Edge
#2::LaunchOrCycle("msedge.exe", "msedge.exe", "edge")

; Win + 3: Microsoft Outlook
#3::LaunchOrCycle("OUTLOOK.EXE", "outlook.exe", "outlook")

; Win + 4: Microsoft Teams
#4::LaunchOrCycle("ms-teams.exe", "ms-teams.exe", "teams")

; Win + 0: Task Manager
#0::LaunchOrCycle("Taskmgr.exe", "taskmgr.exe", "taskmgr")
