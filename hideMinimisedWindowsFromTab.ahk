; AutoHotkey Script zum Minimieren von Fenstern ins System Tray
; Jedes Fenster bekommt sein eigenes Tray-Icon

#Persistent
#SingleInstance Force
SetTitleMatchMode, 2

; Array für minimierte Fenster mit ihren Tray-Icons
global TrayWindows := []
global IconCounter := 1
; Use a slightly decorated header to avoid accidental collisions with window titles
global TrayHeaderText := "• Minimierte Fenster"

; Haupt-Tray-Menu
Menu, Tray, NoStandard
Menu, Tray, Add, Alle Fenster wiederherstellen, RestoreAllWindows
Menu, Tray, Add
Menu, Tray, Add, Hilfe, ShowHelp
Menu, Tray, Add, Beenden, ExitScript
Menu, Tray, Add

; Strg+Alt+H: Aktuelles Fenster ins Tray minimieren
^!h::
    WinGet, ActiveID, ID, A
    if (!ActiveID) {
        TrayTip, Fehler, Kein aktives Fenster gefunden., 2, 3
        return
    }
    
    AppName := GetAppName(ActiveID)

    WinGetTitle, WinTitle, ahk_id %ActiveID%
    
    if (WinTitle = "")
        WinTitle := "Unbenanntes Fenster"
    
    ; Prüfen ob Fenster schon minimiert ist
    for index, WinObj in TrayWindows {
        if (WinObj.ID = ActiveID) {
            TrayTip, Bereits minimiert, Dieses Fenster ist bereits im Tray., 1, 2
            return
        }
    }
    
    ; Fenster verstecken (nicht minimieren!)
    WinHide, ahk_id %ActiveID%
    
    ; Eindeutigen Menu-Namen erstellen
    MenuName := "TrayMenu" . IconCounter
    IconCounter++

    ; Menu für dieses Fenster erstellen
    Menu, %MenuName%, Add, Wiederherstellen, RestoreSingleWindow
    Menu, %MenuName%, Add, Schließen, CloseSingleWindow
    Menu, %MenuName%, Default, Wiederherstellen

    DisplayTitle := AppName . " - " . WinTitle
    ; Kurzen Titel für Tooltip erstellen
    ShortDisplayTitle := DisplayTitle
    if (StrLen(ShortDisplayTitle) > 40)
        ShortDisplayTitle := SubStr(ShortDisplayTitle, 1, 37) . "..."
    Menu, Tray, Add, %ShortDisplayTitle%, :%MenuName%
    
    ; Im Array speichern (store DisplayTitle and icon so deletion is reliable)
    TrayWindows.Push({ID: ActiveID, Title: ShortDisplayTitle, MenuName: MenuName})

    TrayTip, Fenster minimiert, % ShortDisplayTitle . " wurde ins Tray minimiert.", 2, 1
return

; Einzelnes Fenster wiederherstellen
RestoreSingleWindow:
    ; A_ThisMenu enthält den Menu-Namen
    MenuName := A_ThisMenu
    
    ; Finde das Fenster im Array über MenuName
    for index, WinObj in TrayWindows {
        if (WinObj.MenuName = MenuName) {
            RestoreWindowAtIndex(index)
            return
        }
    }

    MsgBox, Debug: MenuName = %MenuName% nicht gefunden
return

; Einzelnes Fenster schließen
CloseSingleWindow:
    ; A_ThisMenu enthält den Menu-Namen
    MenuName := A_ThisMenu
    
    ; Finde das Fenster im Array über MenuName
    for index, WinObj in TrayWindows {
        if (WinObj.MenuName = MenuName) {
            WinID := WinObj.ID
            WinClose, ahk_id %WinID%
            
            ; Tray-Icon entfernen
            RemoveTrayIcon(index)
            TrayWindows.RemoveAt(index)
            return
        }
    }
return

; Tray-Icon entfernen
RemoveTrayIcon(index) {
    global TrayWindows
    
    if (index > TrayWindows.Length() || index < 1) ; index starting from 1
        return
    
    WinObj := TrayWindows[index]
    MenuName := WinObj.MenuName

    ; Menu-Item und zugehöriges Menu löschen (mit Fehlerbehandlung)
    ; The tray menu is rebuilt centrally; only remove the per-window submenu here
    try {
        Menu, %MenuName%, DeleteAll
        WinTitle := WinObj.Title
        Menu, Tray, Delete, %WinTitle%
    }
}

; Hilfsfunktion: Ein einzelnes Fenster per Index wiederherstellen
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

    ; Tray-Icon entfernen und Eintrag löschen
    RemoveTrayIcon(index)
    TrayWindows.RemoveAt(index)
    return true
}

; Alle Fenster wiederherstellen
RestoreAllWindows:
    Count := TrayWindows.Length()
    if (Count = 0) {
        TrayTip, Keine Fenster, Keine Fenster sind im Tray., 2, 1
        return
    }

    ; Rückwärts durchgehen, damit Indizes stimmen beim Löschen
    Loop, %Count% {
        index := Count - A_Index + 1
        RestoreWindowAtIndex(index)
    }

    TrayTip, Wiederhergestellt, Alle Fenster wurden wiederhergestellt., 2, 1
return

; Hilfe anzeigen
ShowHelp:
    MsgBox, 64, Tastenkombinationen,
    (
    Strg+Alt+H - Aktuelles Fenster ins Tray minimieren
    
    Jedes minimierte Fenster bekommt sein eigenes Icon
    im System Tray (unten rechts neben der Uhr).
    
    Rechtsklick auf Fenster-Icon:
    - Wiederherstellen
    - Schließen
    
    Rechtsklick auf Haupt-Icon (AHK):
    - Alle Fenster wiederherstellen
    )
return

; Script beenden
ExitScript:
    ; Alle Fenster wiederherstellen beim Beenden
    Gosub, RestoreAllWindows
    ExitApp
return

GetAppName(WinID) {
    WinGet, ProcessName, ProcessName, ahk_id %WinID%
    if(ProcessName = "")
        return "UnbekannteApp"
    SplitPath, ProcessName, , , , AppName
    
    ; Ersten Buchstaben großschreiben
    StringUpper, FirstChar, AppName, T
    AppName := SubStr(FirstChar, 1, 1) . SubStr(AppName, 2)
    
    return AppName
}