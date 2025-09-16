#Requires AutoHotkey v2.0+
Persistent 1
#SingleInstance Force
IniSettings := IniRead(A_ScriptDir "\settings.ini", "a", "b", "")
CoordMode("Mouse", "Screen")
userConfig := Map()
;-------------------------------------------
; User Config
;-------------------------------------------
; userConfig["LocalUser"] := {shortcut: "shortcut name", title: "RDP window title", id: "", status: ""}
; keep id and status as ""
userConfig["User1"] := {shortcut: "shortcut1", title: "window title 1", id: "", status: ""}
userConfig["User2"] := {shortcut: "shortcut2", title: "window title 2", id: "", status: ""}
userConfig["User3"] := {shortcut: "shortcut3", title: "window title 3", id: "", status: ""}
;-------------------------------------------
; Admin + Registry Check
;-------------------------------------------
SetTitleMatchMode(2)
matchedNumber := 0
if !A_IsAdmin {
    MsgBox("This script must be run as administrator. Exiting script")
    ExitApp
}
try {
    SetRegView 64
    RegRead("HKLM\SOFTWARE\Microsoft\Terminal Server Client", "RemoteDesktop_SuppressWhenMinimized")
} catch {
    MsgBox "RemoteDesktop_SuppressWhenMinimized not found. Command copied to clipboard.`nRun as administrator in cmd. Exiting script"
    A_Clipboard := 'reg add "HKLM\Software\Microsoft\Terminal Server Client" /v RemoteDesktop_SuppressWhenMinimized /t REG_DWORD /d 2'
    ExitApp
}
;-------------------------------------------
; User ID Check (some testing stuff)
;-------------------------------------------
RunWait(A_ComSpec ' /c qwinsta > sessionlist.txt', , 'Hide')
sessionList := FileRead(A_ScriptDir "\sessionList.txt")
for username, data in userConfig {
    if RegExMatch(sessionList, username . "\s+(\d+)\s+(\w+)", &match)
        data.id := match[1], data.status := match[2]
    else {
        data.id := "", data.status := ""
        MsgBox username " not found in session list.`nMake sure it's launched to test this script correctly."
    }
}
FileDelete(A_ScriptDir "\sessionList.txt")
;-------------------------------------------
; GUI creation
;-------------------------------------------
MyGui := Gui("+AlwaysOnTop +ToolWindow +Border +OwnDialogs")
MyGui.Add("Text", "vTextStuff x10 y8 w180 h20 Center", "Debug, expect this to change")
MyGui.Add("Text", "x10 y30 w180 h20 Center", "Reconnect Count: 0")
MyGui.Title := "Press F6 to Close"
MyGui.Show("x" A_ScreenWidth - 205 " y40 w200 h60")

if IniSettings = "True" {
    response := MsgBox("I recommend pressing F8 for troubleshooting. Show this GUI again next time?", "Setup Check", 4)
    if response = "No"
        IniWrite("False", A_ScriptDir "\settings.ini", "a", "b")
}
;-------------------------------------------
; Main Loop
;-------------------------------------------
reconnectCount := 0
loop {
    WinWait("Remote Desktop Connection", , 60)
    RunWait(A_ComSpec ' /c qwinsta > "' A_ScriptDir '\sessionList.txt"', , 'Hide')
    sessionList := FileRead(A_ScriptDir "\sessionList.txt")
    hasDisconnectedOrMissing := false

    for username, data in userConfig {
        if RegExMatch(sessionList, username . "\s+(\d+)\s+(\w+)", &match) {
            data.id := match[1], data.status := match[2]
            if data.status != "Active"
                hasDisconnectedOrMissing := true
        } else {
            data.id := "", data.status := ""
            hasDisconnectedOrMissing := true
        }
    }
    FileDelete(A_ScriptDir "\sessionList.txt")

    if WinExist("Remote Desktop Connection") || hasDisconnectedOrMissing {
        try {
            hwnd := WinExist("Remote Desktop Connection")
            if hwnd {
                pid := WinGetPID(hwnd)
                winList := WinGetList("ahk_pid " pid " ahk_exe mstsc.exe")
                if (hasDisconnectedOrMissing || (winList.Length > 1)) {
                    RunWait(A_ComSpec ' /c taskkill /F /PID ' pid, , 'Hide')
                    WinWaitClose("ahk_pid " pid " ahk_exe mstsc.exe", , 3)
                }
            }
        }
        totalReconnects := 0
        MyGui["Count"].Text := "Reconnect Count: " reconnectCount
        MyGui["TextStuff"].Text := "Attempting to reconnect"

        RunWait(A_ComSpec ' /c qwinsta > "' A_ScriptDir '\sessionList.txt"', , 'Hide')
        sessionList := FileRead(A_ScriptDir "\sessionList.txt")

        for username, data in userConfig {
            if RegExMatch(sessionList, username . "\s+(\d+)\s+(\w+)", &match) {
                data.id := match[1], data.status := match[2]
            } else {
                data.id := "", data.status := ""
            }
        }
        FileDelete(A_ScriptDir "\sessionList.txt")

        for username, data in userConfig {
            if data.status != "Active" {
                if data.id != ""
                    try RunWait(A_ComSpec ' /c logoff ' data.id, , 'Hide')
                Sleep(1000)
                try WinClose(data.title)
                try Run(A_ScriptDir "\" data.shortcut ".lnk", , 'Hide')
                Sleep(2500)
                totalReconnects++
            }
        }

        if totalReconnects > 0 {
            reconnectCount++
            Sleep(7500)
            for username, data in userConfig {
                if WinWait(data.title, , 10) {
                    try WinMove(10, 10)
                    try WinMinimize(data.title)
                }
            }
            MouseMove(1, 1)
            Sleep(1000)
        }

        MyGui["TextStuff"].Text := "waiting for window"
    }
}

;-------------------------------------------
; Hotkeys
;-------------------------------------------
F6:: {
    MyGui.Destroy()
    ExitApp
}
F7:: {
    MyGui.Destroy()
    Reload
}
F8:: {
    Run(A_ComSpec ' /k qwinsta')
    for username, data in userConfig
        MsgBox "User: " username "`n" "ID: " data.id "`n" "Title: " data.title "`n" "Shortcut: " data.shortcut "`n" "Status: " data.status
}
