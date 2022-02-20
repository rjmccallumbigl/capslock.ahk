; Use a popup window to convert time zones and save to clipboard
; Inspiration: https://www.autohotkey.com/boards/viewtopic.php?t=36691

; Build GUI
Gui, Add, DateTime, vMyEdit gformatDates choose%A_NowUTC%, MMM. d @ h:mmtt 'UTC'
Gui, Add, Edit, w200 vDate1
Gui, Add, Edit, w200 vDate2
Gui, Add, Edit, w200 vDate3
Gui, Add, Edit, w200 vDate4
Gui, Add, Button, , &Send All to Clipboard
Gui, Add, Button, Default, &Convert Time
Gui +MinimizeBox
Gui, Show, , Insert Specified Time
return

; Send to clipboard and close
ButtonSendAlltoClipboard(){
    global dateEST
    global datePST
    global dateUTC
    global dateIST
    string := dateEST . " [" . datePST . " | " . dateUTC . " | " . dateIST . "]"
    Clipboard := string
    WinClose, Insert Specified Time
    return
}

; Format date + time
ButtonConvertTime:
formatDates:
    GuiControlGet, MyEdit
    utcTime := MyEdit
    backupString := utcTime
    estOffset := -5
    pstOffset := -8
    istOffset := 5.5
    FormatTime, dateUTC, %utcTime%, MMM. d @ h:mmtt UTC
    GuiControl, Text, Date1, %dateUTC%
    utcTime += estOffset, hours
    FormatTime, dateEST, %utcTime%, MMM. d @ h:mmtt EST
    GuiControl, Text, Date2, %dateEST%
    utcTime := backupString

    utcTime += pstOffset, hours
    FormatTime, datePST, %utcTime%, MMM. d @ h:mmtt PST
    GuiControl, Text, Date3, %datePST%
    utcTime := backupString

    utcTime += istOffset, hours
    FormatTime, dateIST, %utcTime%, MMM. d @ h:mmtt IST
    GuiControl, Text, Date4, %dateIST%
    utcTime := backupString
return

; Exit
GuiEscape:
GuiClose:
ButtonCancel:
ExitApp
