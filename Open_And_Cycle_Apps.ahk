#Requires AutoHotkey v2.0

; Helper function to check if window is cloaked by DWM
IsCloaked(hwnd) {
    static DWMWA_CLOAKED := 14
    flag := 0
    ok := DllCall("dwmapi\DwmGetWindowAttribute"
        , "ptr", hwnd, "int", DWMWA_CLOAKED
        , "int*", &flag, "int", 4, "int")
    return (ok = 0) && (flag != 0)
}

; Helper function to validate if window is real and interactable
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

; Generic function to launch or cycle through application windows
LaunchOrCycle(processName, launchCommand, appKey) {
    ; Static variables for each app (using appKey to differentiate)
    static snapshots := Map()
    
    ; Initialize app snapshot if it doesn't exist
    if (!snapshots.Has(appKey)) {
        snapshots[appKey] := {
            windowSnapshot: [],
            lastActivated: 0,
            snapshotTime: 0
        }
    }
    
    appData := snapshots[appKey]
    
    ; Check if snapshot is expired (5 seconds timeout)
    currentTime := A_TickCount
    if (currentTime - appData.snapshotTime > 5000 || appData.windowSnapshot.Length = 0) {
        ; Take new snapshot
        wins := WinGetList("ahk_exe " . processName)
        appData.windowSnapshot := []
        
        ; Filter for real windows
        for id in wins {
            if (IsRealWindow(id)) {
                appData.windowSnapshot.Push(id)
            }
        }
        
        appData.snapshotTime := currentTime
        appData.lastActivated := 0  ; Reset counter for new snapshot
    }
    
    ; If no windows found, launch the application
    if (appData.windowSnapshot.Length = 0) {
        LaunchAndFocus(launchCommand, processName, appKey)
        return
    }
    
    ; If only one window, just activate it
    if (appData.windowSnapshot.Length = 1) {
        ActivateWindow(appData.windowSnapshot[1])
        return
    }
    
    ; Cycle through the snapshot
    appData.lastActivated := (appData.lastActivated >= appData.windowSnapshot.Length) ? 1 : appData.lastActivated + 1
    ActivateWindow(appData.windowSnapshot[appData.lastActivated])
}

; Helper function to activate window with validation
ActivateWindow(hwnd) {
    if (IsRealWindow(hwnd)) {
        ; Restore if minimized
        if (WinGetMinMax("ahk_id " hwnd) = -1) {
            WinRestore("ahk_id " hwnd)
        }
        Sleep(30)
        WinActivate("ahk_id " hwnd)
        WinWaitActive("ahk_id " hwnd, , 0.15)
    }
}

; Launch application and focus with timer-based polling
LaunchAndFocus(launchCommand, processName, appKey) {
    static activeTimers := Map()
    
    ; Cancel existing timer for this app if any
    if (activeTimers.Has(appKey)) {
        SetTimer(activeTimers[appKey], 0)
        activeTimers.Delete(appKey)
    }
    
    ; Launch the application
    pid := Run(launchCommand)
    launchTime := A_TickCount
    attempts := 0
    
    ; Create polling function
    PollForWindow() {
        attempts++
        
        ; First try PID match
        for hwnd in WinGetList("ahk_pid " pid) {
            if (IsRealWindow(hwnd)) {
                ActivateWindow(hwnd)
                SetTimer(PollForWindow, 0)
                activeTimers.Delete(appKey)
                return
            }
        }
        
        ; After 400ms, fallback to process name match
        if (A_TickCount - launchTime > 400) {
            for hwnd in WinGetList("ahk_exe " processName) {
                if (IsRealWindow(hwnd)) {
                    ActivateWindow(hwnd)
                    SetTimer(PollForWindow, 0)
                    activeTimers.Delete(appKey)
                    return
                }
            }
        }
        
        ; Timeout after 5 seconds
        if (A_TickCount - launchTime > 5000) {
            SetTimer(PollForWindow, 0)
            activeTimers.Delete(appKey)
        }
    }
    
    ; Store timer reference and start polling
    activeTimers[appKey] := PollForWindow
    SetTimer(PollForWindow, 75)
}

; LAlt + 1: Windows Terminal
LAlt & 1::LaunchOrCycle("WindowsTerminal.exe", "wt.exe", "terminal")

; LAlt + 3: Vivaldi
LAlt & 3::LaunchOrCycle("vivaldi.exe", "vivaldi.exe", "vivaldi")

; LAlt + 5: Outlook
LAlt & 5::LaunchOrCycle("OUTLOOK.EXE", "outlook.exe", "outlook")

; LAlt + 6: Teams
LAlt & 6::LaunchOrCycle("ms-teams.exe", "ms-teams.exe", "teams")

; LAlt + 7: Obsidian
LAlt & 7::LaunchOrCycle("Obsidian.exe", "obsidian://", "obsidian")

; LAlt + 8: ProtonMail
LAlt & 8::LaunchOrCycle("Proton Mail.exe", "C:\Users\DagOleBergersenHimle\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Proton\Proton Mail.lnk", "protonmail")