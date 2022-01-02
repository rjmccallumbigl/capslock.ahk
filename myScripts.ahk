#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.

; /* ============== */ 
; /* Text Expansion */
; /* ============== */

; @@ - My @
:*:@@::rjmccallumbigl
Return

; ## - Phone number
:*:##::123-456-7890
Return

; ~~ - Current time (EST)
:*:~~::
	FormatTime, time, % A_Now, ddd_MM-dd-yyyy_hh-mmtt_EST
	send %time%
Return

; @home - Home address
:*:@home::123 Fake Street
Return

; @city - City, State ZIP
:*:@city::Charlotte, NC 28215
Return


; /* ============== */ 
; /*     Macros     */
; /* ============== */

; Double click the right mouse button to start the Snipping Tool
; https://www.autohotkey.com/board/topic/7648-right-mouse-button-added-functionality-doubleclick-it/page-2
~RButton Up::
	If (A_PriorHotKey = A_ThisHotKey and A_TimeSincePriorHotkey < 500) {		
		Send,#+s
	}
Return

; Double click the middle mouse button to paste
~MButton Up::
	If (A_PriorHotKey = A_ThisHotKey and A_TimeSincePriorHotkey < 500) {
		Send,^v
	}
Return

; CTRL + Shift + Spacebar - Toggle always on Top for current window
; https://www.maketecheasier.com/favorite-autohotkey-scripts/
^+SPACE:: Winset, Alwaysontop, , A
Return

; Win + scrollLock - Temporarily suspend AutoHotkey in case of compatibility issues
; https://www.maketecheasier.com/favorite-autohotkey-scripts/
#ScrollLock::Suspend
Return

; Win + Shift+ W - Open Chrome in incognito mode
; https://nickjanetakis.com/blog/remap-and-set-global-hotkeys-on-windows-10-with-auto-hotkey
#+w::Run chrome.exe -incognito
Return

; Scroll over taskbar to change volume, Middle click toggles mute
#If MouseIsOver("ahk_class Shell_TrayWnd")
{
	WheelUp::Send {Volume_Up}
	WheelDown::Send {Volume_Down}
	~MButton Up::Send {Volume_Mute}
}
Return
MouseIsOver(WinTitle) {
	MouseGetPos,,, Win
return WinExist(WinTitle . " ahk_id " . Win)
}
