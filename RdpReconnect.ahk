/************************************************************************
 * @description RDP Reconnect - keeps up your RDP sessions open at all times.
 * @author Dully176
 * @date 2025/09/22
 * @version 0.0.1
 ***********************************************************************/
#Requires AutoHotkey v2.0+
Persistent 1
#SingleInstance Force
IniSettings := IniRead(A_ScriptDir "\settings.ini", "a", "b", "")
CoordMode("Mouse", "Screen")
userConfig := []
; ================
; User Config
; userConfig.Push({user: "WindowsUserName", shortcut: "filename", title: "RDP window title", id: "", status: ""})
; ================
userConfig.Push({user: "User1", shortcut: "shortcut1", title: "WindowTitle1", id: "", status: ""})
userConfig.Push({user: "User2", shortcut: "shortcut2", title: "WindowTitle2", id: "", status: ""})
userConfig.Push({user: "User3", shortcut: "shortcut3", title: "WindowTitle3", id: "", status: ""})
; ================
; Admin + Registry Check
; ================
SetTitleMatchMode(2)
matchedNumber := 0
if !A_IsAdmin {
    MsgBox("This script must be ran as administrator. Exiting script")
    ExitApp
}
try {
    SetRegView 64
    RegRead("HKLM\SOFTWARE\Microsoft\Terminal Server Client", "RemoteDesktop_SuppressWhenMinimized")
} catch {
    MsgBox "RemoteDesktop_SuppressWhenMinimized not found.`nThis is so you can minimize the RDP sessions, which this macro uses.`nCommand copied to clipboard.`nRun as administrator in cmd. Exiting script"
    A_Clipboard := 'reg add "HKLM\Software\Microsoft\Terminal Server Client" /v RemoteDesktop_SuppressWhenMinimized /t REG_DWORD /d 2'
    ExitApp
}
; ================
; User ID Check (some testing stuff)
; ================
UpdateStatus(), temp := ""
for _, data in userConfig {
    if data.status = ""
        temp .= data.user "."
}
(temp? MsgBox(temp " not found in session list.`nMake sure that RDP session is launched."):"")
; ================
; GUI creation
; ================
MyGui := Gui("+AlwaysOnTop +ToolWindow +Border +OwnDialogs")
MyGui.OnEvent("Close", (*) => ExitApp())
MyGui.Add("Text", "vStatusText x10 y8 w180 h20 Center", "Monitoring RDP sessions")
MyGui.Add("Text", "vReconnectCount x10 y30 w180 h20 Center", "Reconnects: 0")
MyGui.Title := "Press F6 to Close"
MyGui.Show("x" A_ScreenWidth - 205 " y40 w200 h60")

if IniSettings = "True" {
    response := MsgBox("I recommend pressing F8 for troubleshooting. Show this GUI again next time?", "Setup Check", 4)
    if response = "No"
        IniWrite("False", A_ScriptDir "\settings.ini", "a", "b")
}
; ================
; Main Loops
; ================
reconnectCount := 0
loop {
    WinWait("Remote Desktop Connection", , 60)
    needsReconnect := UpdateStatus()
    if WinExist("Remote Desktop Connection") || needsReconnect {
        MyGui["StatusText"].Text := "Disconnect Detected"
        while (hwnd := WinExist("Remote Desktop Connection")) {
            try {
                pid := WinGetPID(hwnd)
                RunWait(A_ComSpec ' /c taskkill /F /PID ' pid, , 'Hide')
                WinWaitClose("ahk_id " hwnd, , 3)
            } catch {
                try WinClose("ahk_id " hwnd)
            }
        }
        MyGui["StatusText"].Text := "Attempting to reconnect"
        totalDisconnected := 0
        UpdateStatus()
        for _, data in userConfig {
            if data.status != "Active" {
                if data.id != ""
                    try RunWait(A_ComSpec ' /c logoff ' data.id, , 'Hide')
                Sleep(1000)
                try WinClose(data.title)
                if !FileExist(shortcut := A_ScriptDir "\" data.shortcut ".lnk")
                    continue
                try Run(shortcut, , 'Hide')
                totalDisconnected++
                Sleep(2500)
            }
        }

        if totalDisconnected > 0 {
            reconnectCount++
            MyGui["ReconnectCount"].Text := "Reconnects: " reconnectCount
            Sleep(7500)
            MyGui["StatusText"].Text := "Minimizing RDP Sessions"
            for _, data in userConfig {
                if WinWait(data.title, , 10) {
                    try WinMove(10, 10)
                    try WinMinimize(data.title)
                }
            }
            MouseMove(1, 1)
            Sleep(1000)
        }

        MyGui["StatusText"].Text := "Monitoring RDP sessions"
    }
}
; ================
; Functions
; ================
UpdateStatus() {
    global userConfig
    RunWait(A_ComSpec ' /c qwinsta > "' A_ScriptDir '\sessionList.txt"', , 'Hide')
    sessionList := FileRead(A_ScriptDir '\sessionList.txt'), needsReconnect := false
    for _, data in userConfig {
        if RegExMatch(sessionList, data.user . "\s+(\d+)\s+(\w+)", &match) {
            data.id := match[1]
            data.status := match[2]
            if (data.status != "Active")
                needsReconnect := true
        } else {
            data.id := "", data.status := ""
            needsReconnect := true
        }
    }
    FileDelete(A_ScriptDir "\sessionList.txt")
    return needsReconnect
}
; ================
; Hotkeys
; ================
F6:: {
    try MyGui.Destroy()
    ExitApp
}
F7:: {
    try MyGui.Destroy()
    Reload
}
F8:: {
    Run(A_ComSpec ' /k qwinsta')
    UpdateStatus()
    temp := ""
    for _, data in userConfig
        temp .= "User: " data.user " | Shortcut: " data.shortcut " | Title: " data.title "`nID: " data.id " | Status: " data.status "`n`n"
    MsgBox(temp, "Current Status")
}
