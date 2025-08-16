#Requires AutoHotkey v2.0

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
        
        ; Filter for visible windows
        for id in wins {
            minMax := WinGetMinMax("ahk_id " id)
            if (minMax != -1) {
                appData.windowSnapshot.Push(id)
            }
        }
        
        appData.snapshotTime := currentTime
        appData.lastActivated := 0  ; Reset counter for new snapshot
    }
    
    ; If no windows found, launch the application
    if (appData.windowSnapshot.Length = 0) {
        Run(launchCommand)
        return
    }
    
    ; If only one window, just activate it
    if (appData.windowSnapshot.Length = 1) {
        WinActivate("ahk_id " appData.windowSnapshot[1])
        return
    }
    
    ; Cycle through the snapshot
    appData.lastActivated := (appData.lastActivated >= appData.windowSnapshot.Length) ? 1 : appData.lastActivated + 1
    WinActivate("ahk_id " appData.windowSnapshot[appData.lastActivated])
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