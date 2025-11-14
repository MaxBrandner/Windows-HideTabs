; Icon: prefer the script directory so the icon is found when the script is moved
I_Icon := A_ScriptDir . "\\menu.ico"
; <a href="https://www.flaticon.com/free-icons/app-menu" title="app menu icons">App menu icons created by alkhalifi design - Flaticon</a>
IfExist, %I_Icon%
Menu, Tray, Icon, %I_Icon%
;return

; AutoHotkey Script to minimize windows to system tray

#Persistent
#SingleInstance Force
SetTitleMatchMode, 2

; Array for minimized windows with their tray icons
global TrayWindows := []
global MinimizedAppsCounter := 1
; Use a slightly decorated header to avoid accidental collisions with window titles

; Main tray menu
Menu, Tray, NoStandard
Menu, Tray, Add, Restore All Windows, RestoreAllWindows
Menu, Tray, Add
Menu, Tray, Add, Help, ShowHelp
Menu, Tray, Add, Exit, ExitScript
Menu, Tray, Add

; Ctrl+Alt+H: Minimize current window to tray
^!h::
    WinGet, ActiveID, ID, A
    if (!ActiveID) {
        TrayTip, Error, No active window found., 2, 3
        return
    }
    
    AppName := GetAppName(ActiveID)
    ExePath := GetExePath(ActiveID)

    WinGetTitle, WinTitle, ahk_id %ActiveID%
    
    if (WinTitle = "")
        WinTitle := "Unknown Window"
    
    ; Check if window is already minimized
    for index, WinObj in TrayWindows {
        if (WinObj.ID = ActiveID) {
            TrayTip, Already Minimized, This window is already in the tray., 1, 2
            return
        }
    }
    
    ; Hide window
    WinHide, ahk_id %ActiveID%
    
    ; Create unique menu name
    MenuName := "TrayMenu" . MinimizedAppsCounter
    MinimizedAppsCounter++

    ; Create menu for this window
    Menu, %MenuName%, Add, Restore, RestoreSingleWindow
    Menu, %MenuName%, Add, Close, CloseSingleWindow
    Menu, %MenuName%, Default, Restore

    DisplayTitle := AppName . " - " . WinTitle
    ; Create tooltip title
    ShortDisplayTitle := DisplayTitle
    if (StrLen(ShortDisplayTitle) > 40)
        ShortDisplayTitle := SubStr(ShortDisplayTitle, 1, 37) . "..."
    Menu, Tray, Add, %ShortDisplayTitle%, :%MenuName%
    
    ; Store in array 
    TrayWindows.Push({ID: ActiveID, Title: ShortDisplayTitle, MenuName: MenuName})

    TrayTip, Window Minimized, % ShortDisplayTitle . " has been minimized to tray.", 2, 1

    ; Set icon (try to extract icon from executable)
    if (ExePath != "") {
        try {
            Menu, Tray, Icon, %ShortDisplayTitle%, %ExePath%, , 16
            return
        }
    }
    ; Fallback: Standard icon
    Menu, Tray, Icon, %ShortDisplayTitle%, shell32.dll, 3, 16
return

; Restore single window
RestoreSingleWindow:
    ; A_ThisMenu contains the menu name
    MenuName := A_ThisMenu
    
    ; Find window in array by MenuName
    for index, WinObj in TrayWindows {
        if (WinObj.MenuName = MenuName) {
            RestoreWindowAtIndex(index)
            return
        }
    }

    MsgBox, Debug: MenuName = %MenuName% not found
return

; Close single window
CloseSingleWindow:
    ; A_ThisMenu contains the menu name
    MenuName := A_ThisMenu
    
    ; Finde das Fenster im Array über MenuName
    for index, WinObj in TrayWindows {
        if (WinObj.MenuName = MenuName) {
            WinID := WinObj.ID
            WinClose, ahk_id %WinID%
            
            ; Remove app from tray menu and delete entry
            RemoveAppFromTrayMenu(index)
            TrayWindows.RemoveAt(index)
            return
        }
    }
return

; Remove app from tray menu
RemoveAppFromTrayMenu(index) {
    global TrayWindows
    
    if (index > TrayWindows.Length() || index < 1) ; index starting from 1
        return
    
    WinObj := TrayWindows[index]
    MenuName := WinObj.MenuName

    ; Delete menu item and associated menu (with error handling)
    ; The tray menu is rebuilt centrally; only remove the per-window submenu here
    try {
        Menu, %MenuName%, DeleteAll
        WinTitle := WinObj.Title
        Menu, Tray, Delete, %WinTitle%
    }
}

; Helper function: Restore single window by index
RestoreWindowAtIndex(index) {
    global TrayWindows

    if (index < 1 || index > TrayWindows.Length())
        return false

    WinObj := TrayWindows[index]
    WinID := WinObj.ID

    WinShow, ahk_id %WinID%
    Sleep, 100
    WinRestore, ahk_id %WinID%
    Sleep, 100
    WinActivate, ahk_id %WinID%

    ; Remove app from tray menu and delete entry
    RemoveAppFromTrayMenu(index)
    TrayWindows.RemoveAt(index)
    return true
}

; Restore all windows
RestoreAllWindows:
    Count := TrayWindows.Length()
    if (Count = 0) {
        TrayTip, No Windows, No windows are in the tray., 2, 1
        return
    }

    ; Iterate backwards so indices remain correct when deleting
    Loop, %Count% {
        index := Count - A_Index + 1
        RestoreWindowAtIndex(index)
    }

    TrayTip, Restored, All windows have been restored., 2, 1
return

; Show help
ShowHelp:
    MsgBox, 64, Keyboard Shortcuts,
    (
    Ctrl+Alt+H - Minimize current window to tray
    
    Each minimized window gets its own menu entry
    
    Right-click on main icon (AHK):
    - Restore All Windows
    ---
    - Help (this message)
    - Exit
    ---
    - List of all minimized windows

    Right-click on window entry:
    - Restore
    - Close   
    )
return

; Exit script
ExitScript:
    ; Restore all windows before exiting
    Gosub, RestoreAllWindows
    ExitApp
return

GetAppName(WinID) {
    WinGet, ProcessName, ProcessName, ahk_id %WinID%
    if(ProcessName = "")
        return "UnknownApp"
    SplitPath, ProcessName, , , , AppName
    
    ; Capitalize first letter
    StringUpper, FirstChar, AppName, T
    AppName := SubStr(FirstChar, 1, 1) . SubStr(AppName, 2)
    
    return AppName
}

; Helper function: Get full executable path
GetExePath(WinID) {
    WinGet, ProcessPath, ProcessPath, ahk_id %WinID%
    return ProcessPath
}