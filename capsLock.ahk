#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
#include <Vis2>

/*
**********************************************************************************
 *
 * DESCRIPTION
 * Slows down and extends the CAPSLOCK key. Based on CAPSHIFT.
 *
 * SOURCE
 * https://www.autohotkey.com/board/topic/4310-capshift-slow-down-and-extend-the-caps-lock-key/
 *
 * VERSION
 * 0.7.3
 *
 * TODO
 *  Replace specific text with another string in highlighted text
 *	Storage Usage
 *	Indicate what processes/programs are using this file (handles)
 *
**********************************************************************************
*/

; Declare variables
SetTimer, TOOLTIP, 1500
SetTimer, TOOLTIP, Off
TimeCapsToggle = 5
TimeOut = 30
global stringGlobal := ""
global MouseX
global MouseY

; Build GUI in case we use the Insert Specified Time function
Gui, Add, DateTime, vutcTime gformatDates choose%A_NowUTC%, MMM. d @ h:mmtt 'UTC'
Gui, Add, Edit, w200 vDate1
Gui, Add, Edit, w200 vDate2
Gui, Add, Edit, w200 vDate3
Gui, Add, Edit, w200 vDate4
Gui, Add, Edit, w200 vDate5
Gui, Add, Button, Default, &OK
Gui +MinimizeBox

; Enable mousewheel in AutoHotkey GUIs
#If MouseIsOver("ahk_class AutoHotkeyGUI")
	WheelUp::Send {Up}
	WheelDown::Send {Down}
#If
MouseIsOver(WinTitle){
	MouseGetPos, , , Win
	Return WinExist(WinTitle . " ahk_id " . Win)
}

; Define hotstring, trigger via hold CapsLock for 1.5s or hold Right Click + Middle click
CapsLock::
~RButton & MButton::
	OldClipboard:= ClipboardAll ;Save existing clipboard.
	counter = 0
	MouseGetPos, MouseX, MouseY
	Progress, ZH16 ZX0 ZY0 B R0-%TimeOut%
	Loop, %TimeOut%
	{
		Sleep, 10
		counter += 1
		Progress, %counter%
		If (counter = TimeCapsToggle)
			Progress, ZH16 ZX0 ZY0 B R0-%TimeOut% CBFF0000
		if (A_ThisHotkey = "CapsLock")
		{
			GetKeyState, state, CapsLock, P
			If state = U
				Break
		}

		Else If (A_ThisHotkey = "MButton")
		{
			GetKeyState, state, MButton, P
			If state = U
				Break
		}
	}
	Progress, Off
	If counter = %TimeOut%
		Gosub, MENU
	Else If (counter > TimeCapsToggle)
	{
		Gosub, CapsLock_State_Toggle
	}

	if (stringGlobal){
		clipboard := stringGlobal
	} else {
		Clipboard := OldClipboard
	}
Return

; Build CapsLock menu
MENU:
Winget, Active_Window, ID, A
Send, ^c
ClipWait, 1
; Default string manipulation
Menu, convert, Add
Menu, convert, Delete
Menu, convert, Add, &CapsLock Toggle, CapsLock_State_Toggle
If (GetKeyState("CapsLock", "T") = True)
	Menu, Convert, Check, &CapsLock Toggle
Else
	Menu, Convert, uncheck, &CapsLock Toggle
Menu, convert, Add,
Menu, convert, Add, &UPPER CASE, MENU_ACTION
Menu, convert, Add, &lower case, MENU_ACTION
Menu, convert, Add,
Menu, convert, Add, &Title Case, MENU_ACTION
Menu, convert, Add, &Sentence case, MENU_ACTION
Menu, convert, Add, &iNVERT cASE, MENU_ACTION
Menu, convert, Add, &SpOnGeBoB, MENU_ACTION
Menu, convert, Add, &S p r e a d T e x t, MENU_ACTION
Menu, convert, Add,
; Advanced string manipulation
Menu, Modify Text..., Add, Remove_&under_scores, MENU_ACTION
Menu, Modify Text..., Add, Remove.&full.stops, MENU_ACTION
Menu, Modify Text..., Add, Remove-&dashes, MENU_ACTION
Menu, Modify Text..., Add, Remove_illegal_characters, MENU_ACTION
Menu, Modify Text..., Add, Remove &emojis, MENU_ACTION
Menu, Modify Text..., Add, Remove &Uppercase, MENU_ACTION
Menu, Modify Text..., Add, Remove &Lowercase, MENU_ACTION
Menu, Modify Text..., Add, &`Sort, MENU_ACTION
Menu, Modify Text..., Add, &snake_Case to camelCase, MENU_ACTION
Menu, Modify Text..., Add, &Swap at Anchor Word, MENU_ACTION
Menu, Modify Text..., Add, &Tabs to Spaces, MENU_ACTION
Menu, Modify Text..., Add, &Spaces to Tabs, MENU_ACTION
Menu, convert, Add, &Modify Text..., :Modify Text...
Menu, convert, Add,
; Wrap text (coding)
Menu, Wrap Text..., Add, &`(...), MENU_ACTION
Menu, Wrap Text..., Add, &`{...}, MENU_ACTION
Menu, Wrap Text..., Add, &`[...], MENU_ACTION
Menu, Wrap Text..., Add, &`<...>, MENU_ACTION
Menu, Wrap Text..., Add, `<&#...#>, MENU_ACTION
Menu, Wrap Text..., Add, &`/*...*/, MENU_ACTION
Menu, Wrap Text..., Add, &`'...', MENU_ACTION
Menu, Wrap Text..., Add, &`"...", MENU_ACTION
Menu, convert, Add, &Wrap Text..., :Wrap Text...
Menu, convert, Add,
; Utilize internet browser
Menu, Browser Search..., Add, &Google, MENU_ACTION
Menu, Browser Search..., Add, &Thesaurus, MENU_ACTION
Menu, Browser Search..., Add, &Wikipedia, MENU_ACTION
Menu, Browser Search..., Add, &Define, MENU_ACTION
Menu, Browser Search..., Add, &Open Page..., MENU_ACTION
Menu, Browser Search..., Add, &Translate..., MENU_ACTION
Menu, convert, Add, &Browser Search..., :Browser Search...
Menu, convert, Add,
; Utilize Windows Explorer
Menu, Explorer..., Add, &Open Folder..., MENU_ACTION
Menu, Explorer..., Add, &Open Parent Folder..., MENU_ACTION
Menu, Explorer..., Add, &Copy File Names and Details from Folder to Clipboard..., MENU_ACTION
Menu, Explorer..., Add, &Add File to Subfolder..., MENU_ACTION
Menu, Explorer..., Add, &Pull File Out to Parent Folder..., MENU_ACTION
Menu, Explorer..., Add, &Clipboard to File..., MENU_ACTION
Menu, Explorer..., Add, &File to Clipboard..., MENU_ACTION
Menu, convert, Add, &Explorer..., :Explorer...
Menu, convert, Add,
; Run scripts
Menu, Scripts..., Add, &Run in CMD, MENU_ACTION
Menu, Scripts..., Add, &Run in PowerShell, MENU_ACTION
Menu, Scripts..., Add, &Run in Ubuntu, MENU_ACTION
Menu, Scripts..., Add, Check Sensitivity with &Alex.ps1 (npm), MENU_ACTION
Menu, Scripts..., Add, &BleachBit, MENU_ACTION
Menu, Scripts..., Add, &RemoveResourceGroups.ps1, MENU_ACTION
Menu, Scripts..., Add, &createNewVM.ps1, MENU_ACTION
Menu, Scripts..., Add, &{Update NSGS}.ps1, MENU_ACTION
Menu, Scripts..., Add, &getDiskSpace.ps1, MENU_ACTION
Menu, Scripts..., Add, &grepFolder.ps1, MENU_ACTION
Menu, convert, Add, &Scripts..., :Scripts...
Menu, convert, Add,
; Image manipulation
Menu, Images..., Add, Save image from clipboard to &folder, MENU_ACTION
Menu, Images..., Add, OCR with &Vis2.ahk, MENU_ACTION
Menu, Images..., Add, Get &HEX value from cursor, MENU_ACTION
Menu, convert, Add, &Images..., :Images...
Menu, convert, Add,
; Insert/modify datetime strings
Menu, TimeDate, Add, Time/Date, MENU_ACTION
Menu, TimeDate, DeleteAll
List := DateFormats(A_Now)
TextMenuDate(List, "TimeDate", "DateAction")
Menu, convert, Add, &Insert Time/Date..., :TimeDate
Menu, convert, Add, &Insert Specified Time, MENU_ACTION
Menu, convert, Default, &CapsLock Toggle
Menu, convert, Show
Return

; Pass copied string to specified menu action, then paste if required
MENU_ACTION:
	AutoTrim, Off
	string = %clipboard%
	clipboard := Menu_Action(A_ThisMenuItem, string)
	WinActivate, ahk_id %Active_Window%
	If (!stringGlobal){
		; Send, ^v
		SendInput %Clipboard%
	}
	ToolTip, %A_ThisMenuItem%
	SetTimer, TOOLTIP, On
Return

; Define menu actions
Menu_Action(ThisMenuItem, string)
{
	; Convert to UPPER CASE
	If ThisMenuItem =&UPPER CASE
		StringUpper, string, string

	; Convert to lower case
	Else If ThisMenuItem =&lower case
		StringLower, string, string

	; Convert to Title Case
	Else If ThisMenuItem =&Title Case
		StringLower, string, string, T

	; Convert to HaLf CaPs CaSe (AKA alternating caps)
	Else If ThisMenuItem =&SpOnGeBoB
	{
		StringCaseSense, On
		newString := ""
		splitString:= StrSplit(string)
		for index, element in splitString
		{
			If (Mod(index, 2) = 0)
			{
				StringUpper, newChar, element
				newString .= newChar
			} Else {
				StringLower, newChar, element
				newString .= newChar
			}
		}
		string := newString
		StringCaseSense, Off
	}

	; Add extra spaces between every character except spaces and punctuation
	Else If ThisMenuItem =&S p r e a d T e x t
	{
		newString := ""
		splitString:= StrSplit(string)
		for index, element in splitString
		{
			If (element != " "){
				newString .= element
			}
			; Do not space out punctuation. Loooks better to me.
			If !RegExMatch(splitString[index + 1], "[[:punct:]]")
				newString .= " "
		}
		string := newString
	}

	; Convert to iNVERT cASE
	Else If ThisMenuItem =&iNVERT cASE
	{
		StringCaseSense, On
		lower = abcdefghijklmnopqrstuvwxyz
		upper = ABCDEFGHIJKLMNOPQRSTUVWXYZ
		StringLen, length, string
		Loop, %length%
		{
			StringLeft, char, string, 1
			StringGetPos, pos, lower, %char%
			pos += 1
			If pos <> 0
				StringMid, char, upper, %pos%, 1
			Else
			{
				StringGetPos, pos, upper, %char%
				pos += 1
				If pos <> 0
					StringMid, char, lower, %pos%, 1
			}
			StringTrimLeft, string, string, 1
			string .= char
		}
		StringCaseSense, Off
	}

	; Convert to Sentence case
	Else If ThisMenuItem =&Sentence case
	{
		StringCaseSense, On
		lower = abcdefghijklmnopqrstuvwxyz
		upper = ABCDEFGHIJKLMNOPQRSTUVWXYZ
		dot = 1
		StringLen, length, string
		Loop, %length%
		{
			StringLeft, char, string, 1
			if (char == ".")
			{
				dot = 1
			}
			else
			{
				if (dot = 1)
				{
					StringGetPos, pos, lower, %char%
					pos += 1
					If pos <> 0
						StringMid, char, upper, %pos%, 1
				}
				else
				{
					StringGetPos, pos, upper, %char%
					pos += 1
					If pos <> 0
						StringMid, char, lower, %pos%, 1
				}

				if (char <> " ")
					dot = 0
			}
			StringTrimLeft, string, string, 1
			string .= char
		}
		StringCaseSense, Off
	}

	; Remove under_scores
	Else If ThisMenuItem =Remove_&under_scores
		StringReplace, string, string, _, %A_Space%, All

	; Remove periods.from.text
	Else If ThisMenuItem =Remove.&full.stops
		StringReplace, string, string, ., %A_Space%, All

	; Remove all-dashes
	Else If ThisMenuItem =Remove-&dashes
		StringReplace, string, string, -, %A_Space%, All

	; https://www.autohotkey.com/boards/viewtopic.php?f=6&t=69889&p=301478#p301478
	; Modify string to remove illegal chars
	Else If ThisMenuItem =Remove_Illegal_characters
	{
		StringLower, string, string
		string := RegExReplace(string, "[/\\?%*:|<>]", "$u1")
	}
	; Modify string to remove emojis
	Else If ThisMenuItem =Remove &emojis
	{
		; string := RegExReplace(string, "[\uD800-\uDFFF]","")
		string := RegExReplace(string, "\p{C}","")
	}

	; Return lowercase chars
	Else If ThisMenuItem =Remove &Uppercase
		string := RegExReplace(string, "[A-Z]","")

	; Return uppercase chars
	Else If ThisMenuItem =Remove &Lowercase
		string := RegExReplace(string, "[a-z]","")

	; Sort the string
	Else If ThisMenuItem =&`Sort
		Sort, string

	; Convert phrase separated by_or-to lowerCamelCase
	Else If ThisMenuItem =&snake_Case to CamelCase
	{
		string := RegExReplace(string, "(([A-Z]+)|(?i)((?<=[a-z])|[a-z])([a-z]*))[ _-]([a-z]|[A-Z]+)", "$L2$L3$4$T5")
	}

	; Wrap text with parentheses
	Else If ThisMenuItem =&`(...)
		string = (%string%)

	; Wrap text with curly braces
	Else If ThisMenuItem =&`{...}
		string = {%string%}

	; Wrap text with square brackets
	Else If ThisMenuItem =&`[...]
		string = [%string%]

	; Wrap text with angle brackets
	Else If ThisMenuItem =&`<...>
		string = <%string%>

	; Wrap text with angle brackets + octothorpes (block comment)
	Else If ThisMenuItem =`<&#...#>
		string = <# %string% #>

	; Wrap text with slashes + octothorpes (block comment)
	Else If ThisMenuItem =&`/*...*/
		string = /* %string% */

	; Wrap text with single quotes
	Else If ThisMenuItem =&`'...'
		string = '%string%'

	; Wrap text with double quotes
	Else If ThisMenuItem =&`"..."
		string = `"%string%`"

	; Google the highlighted word
	Else If ThisMenuItem =&Google
		Run, https://www.google.com/search?q=%string%

	; Looks the highlighted word up in a thesaurus
	Else If ThisMenuItem =&Thesaurus
		Run, https://www.thesaurus.com/browse/%string%

	; Looks the highlighted word up in Wiki
	Else If ThisMenuItem =&Wikipedia
		Run, https://en.wikipedia.org/wiki/%string%

	; Defines the highlighted word
	Else If ThisMenuItem =&Define
		Run, https://www.google.com/search?q=define+%string%

	; If the copied text is a valid folder, open it in Windows Explorer
	Else If ThisMenuItem =&Open Folder...
	{
		string := StrReplace(string, "/", "\")
		if (FileExist(string))
			Run explorer.exe /select`,%string%
		else
		{
			replacement := RegExReplace(string, "[^\\\/]+[\\\/]?$")
			length := StrLen(replacement)
			subString := SubStr(replacement, 1, length - 1)
			Run explorer.exe %subString%
		}
	}

	; If the copied text is a valid file, open the parent folder in Windows Explorer and highlight the file
	Else If ThisMenuItem =&Open Parent Folder...
	{
		SplitPath, string, , dir
		Run explorer.exe /select`,%dir%
	}

	; If the copied text is a valid site, open it in web browser
	Else If ThisMenuItem =&Open Page...
		Run, chrome.exe %string%

	; Translate in Chrome
	Else If ThisMenuItem =&Translate...
	{
		encodedString := StrReplace(string, A_Space, "+")
		Run, chrome.exe https://www.google.com/search?q=translate+%encodedString%to+english

	}

	; Run  BleachBit portable
	Else If ThisMenuItem =&Bleachbit
		Run, powershell -NoExit -Command "D:\Dropbox\code\bleachbit.ps1"

	; bleachbit.ps1
	;###########################
	; # Cleans system with bleachbit

	; # Declare variables
	; $bleachbit = "D:\Dropbox\Apps\BleachBit-4.4.2-portable\BleachBit-Portable\bleachbit_console.exe"

	; # Run app
	; & $bleachbit --clean windows_explorer.* vlc.* system.recycle_bin system.tmp

	;###########################

	; Use CMD to save open folder contents to clipboard
	Else If (ThisMenuItem ="&Copy File Names and Details from Folder to Clipboard...") {
		; https://lexikos.github.io/v2/docs/FAQ.htm#output
		SplitPath, string, , dir
		stringGlobalTemp := dir . "\stringGlobal.txt"
		RunWait %comspec% ' "/c dir "%dir%" /a /o /q /t > "%stringGlobalTemp%" "', , hide
		stringGlobalFile := FileOpen(stringGlobalTemp, "rw")
		stringGlobal := stringGlobalFile.Read()
		stringGlobalFile.Close()
		FileDelete, %stringGlobalTemp%
	}

	; Save file contents to clipboard
	Else If (ThisMenuItem ="&File to Clipboard..."){
		file := FileOpen(string, "r")
		stringGlobal := File.Read()
		file.Close()
	}

	; Run grepFolder.ps1 script
	Else If ThisMenuItem =&grepFolder.ps1
	{

		; grep.ps1
		;###########################
		; 		# Greps a string after specifying a path to one or more locations. Wildcards are permitted.

		; param (
		; 	[Parameter(Mandatory = $true,
		; 		ValueFromPipeline = $true,
		; 		ValueFromPipelineByPropertyName = $true,
		; 		HelpMessage = "Path to one or more locations.")]
		; 	[ValidateNotNullOrEmpty()]
		; 	[SupportsWildcards()]
		; 	[string[]]
		; 	$path
		; )

		; $pattern = Read-Host -Prompt "Enter query:"
		; Write-Output "Path: $($path)"
		; Write-Output "Pattern: $($pattern)"
		; Select-String -Pattern $pattern -Path "$($path)\*" -Include *.*

		;###########################

		WinGet, WinID, ID, ahk_class CabinetWClass
		CurrentPath := ExplorerPath(WinID)
		Run, powershell -NoExit -Command "D:\Dropbox\code\grep.ps1 -path '%CurrentPath%'"
		stringGlobal := string
	}

	; Run getDiskSpace.ps1 script
	Else If ThisMenuItem =&getDiskSpace.ps1
	{
		; getDiskSpace.ps1
		;###########################
		; # Get disk space
		; Get-CimInstance -ClassName Win32_LogicalDisk

		; Get-CimInstance -Class CIM_LogicalDisk | Select-Object @{Name = "Size(GB)"; Expression = { $_.size / 1gb } }, @{Name = "Free Space(GB)"; Expression = { $_.freespace / 1gb } }, @{Name = "Free (%)"; Expression = { "{0,6:P0}" -f (($_.freespace / 1gb) / ($_.size / 1gb)) } }, DeviceID, DriveType | Where-Object DriveType -EQ '3'
		;###########################
		Run, powershell -NoExit -File "D:\Dropbox\code\getDiskSpace.ps1"
		stringGlobal := string
	}

	; Run RemoveResourceGroups.ps1 script
	Else If ThisMenuItem =&RemoveResourceGroups.ps1
	{
		; RemoveResourceGroups.ps1
		;###########################
		; https://github.com/rjmccallumbigl/Powershell/blob/master/Remove-ResourceGroupsAsyncWithPattern.ps1
		;###########################
		Run, powershell -NoExit -File "C:\Users\rymccall\OneDrive - Microsoft\PowerShell\byronbayer\Powershell\Remove-ResourceGroupsAsyncWithPattern.ps1"
		stringGlobal := string
	}

	; Run createNewVM.ps1 script
	Else If ThisMenuItem =&createNewVM.ps1
	{
		; createNewVM.ps1
		;###########################
		; https://github.com/rjmccallumbigl/Azure-PowerShell---Create-New-VM
		;###########################
		Run, C:\Users\rymccall\GitHub\Azure-PowerShell---Create-New-VM\createNewVM.ps1
		stringGlobal := string
	}

	; Run {Update NSGS}.ps1 script
	Else If ThisMenuItem =&{Update NSGS}.ps1
	{
		; Update your NSGs with a rule scoped to your IP address.ps1
		;###########################
		; https://github.com/rjmccallumbigl/Azure-PowerShell---Add-RDP-rule-to-your-NSGs-with-a-rule-scoped-to-your-IP-address
		;###########################
		Run, powershell -NoExit -File "C:\Users\rymccall\OneDrive - Microsoft\PowerShell\Update your NSGs with a rule scoped to your IP address.ps1"
		stringGlobal := string
	}

	; Run in CMD
	Else If ThisMenuItem =&Run in CMD
	{
		Run, powershell -NoExit -Command "cmd /k 'echo %string% && echo ================================================ && %string%'"
		stringGlobal := string
	}

	; Run in PowerShell
	Else If ThisMenuItem =&Run in PowerShell
	{
		Run, powershell -NoExit -Command "echo '%string%'; `r`n echo ================================================; `r`n %string%"
		stringGlobal := string
	}

	; Run in Ubuntu (WSL)
	Else If ThisMenuItem =&Run in Ubuntu
	{
		Run, powershell -NoExit -Command "echo '%string%'; `r`n echo ================================================; `r`n wsl %string%"
		stringGlobal := string
	}

	; Run Alex to check string for insensitivity (requires Node support)
	; https://github.com/get-alex/alex
	Else If ThisMenuItem =Check Sensitivity with &Alex.ps1 (npm)
	{
		Run, powershell -NoExit -Command """%string%""" | npx alex --stdin"
		stringGlobal := string
	}
	; Type what the anchor word is and swap the text on the opposite sides
	Else If ThisMenuItem =&Swap at Anchor Word
	{
		string := text_swap(string)
	}

	; Convert Tabs to Spaces
	Else If ThisMenuItem =&Tabs to Spaces
	{
		string := TabsToSpaces(string)
	}

	; Convert Spaces to Tabs
	Else If ThisMenuItem =&Spaces to Tabs
	{
		string := RegExReplace(string, "[\s]{4}", A_Tab)
		; string := SpacesToTabs(string)
	}

	; Create a new subfolder, add hightlighted file to it
	Else If ThisMenuItem =&Add File to Subfolder...
	{
		SplitPath, string, , dir, , nameNoExt
		FileCreateDir, %nameNoExt%
		FileMove, %string%, %nameNoExt%
	}

	; Take file out of current folder and add to parent folder
	Else If ThisMenuItem =&Pull File Out to Parent Folder...
	{
		SplitPath, string, , dir
		SplitPath, dir, , dir2
		FileMove, %string%, %dir2%
	}

	; Copy clipboard contents to a text file in a folder currently open in Windows Explorer
	Else If ThisMenuItem =&Clipboard to File...
	{
		WinGet, WinID, ID, ahk_class CabinetWClass
		CurrentPath := ExplorerPath(WinID)
		newPath := CurrentPath . "\clipboardToFile.txt"
		; FileAppend, string`n, %CurrentPath%%newPath%
		file := FileOpen(newPath, "rw")
		file.Write(string "`n")
		file.Close()
	}

	; Convert time and paste string in different time zones
	Else If ThisMenuItem =&Insert Specified Time
	{
		Gui, +LastFound
		Gui, Show, , Insert Specified Time
		GuiHWND := WinExist()
		WinWaitClose, ahk_id %GuiHWND%
		string := ButtonSendAlltoClipboard()
	}

	; Save image from clipboard to folder
	Else If ThisMenuItem =Save image from clipboard to &folder
	{
		WinGet, WinID, ID, ahk_class CabinetWClass
		CurrentPath := ExplorerPath(WinID)
		FormatTime, timeOutputVar, %Date%, ddd_MM-dd-yyyy_hh-mm-sstt_EST
		If (!OldClipboard)
		{
			target := string
		} else {
			target := OldClipboard
		}
		stringGlobal := ImagePutFile(target, CurrentPath . "\" . timeOutputVar . ".png")
	}

	; Use OCR with Vis2 https://github.com/iseahound/Vis2
	Else If ThisMenuItem =OCR with &Vis2.ahk
	{
		stringGlobal := OCR()
	}

	; https://dev.to/radualexandrub/top-8-macros-for-developers-to-maximize-their-productivity-with-ahk-5bfj
	Else If ThisMenuItem =Get &HEX value from cursor
	{
		PixelGetColor, color, %MouseX%, %MouseY%, RGB
		StringLower, color, color
		stringGlobal := "#" . SubStr(color, 3)
	}

	Return string
}

; Return one line time conversion string
ButtonSendAlltoClipboard(){
	global dateEST
	global datePST
	global dateCST
	global dateUTC
	global dateIST
	string := dateEST . " [" . datePST . " | " . dateCST . " | " . dateUTC . " | " . dateIST . "]"
	return string
}

; Close GUI when pressing OK
ButtonOK:
	WinClose, Insert Specified Time
return

; Format date for different time zones
formatDates:
	GuiControlGet, utcTime
	backupString := utcTime
	estObject := time("America/New_York")
	pstObject := time("PST8PDT")
	cstObject := time("CST6CDT")
	estOffset := estObject.utc_offset
	pstOffset := pstObject.utc_offset
	cstOffset := cstObject.utc_offset
	istOffset := 5.5

	; Format UTC
	FormatTime, dateUTC, %utcTime%, MMM. d @ h:mmtt UTC
	GuiControl, Text, Date1, %dateUTC%

	; Format EST
	utcTime += estOffset, hours
	FormatTime, dateEST, %utcTime%, MMM. d @ h:mmtt EST
	GuiControl, Text, Date2, %dateEST%
	utcTime := backupString

	; Format PST
	utcTime += pstOffset, hours
	FormatTime, datePST, %utcTime%, MMM. d @ h:mmtt PST
	GuiControl, Text, Date3, %datePST%
	utcTime := backupString

	; Format CST
	utcTime += cstOffset, hours
	FormatTime, dateCST, %utcTime%, MMM. d @ h:mmtt CST
	GuiControl, Text, Date4, %dateCST%
	utcTime := backupString

	; Format IST
	utcTime += istOffset, hours
	FormatTime, dateIST, %utcTime%, MMM. d @ h:mmtt IST
	GuiControl, Text, Date5, %dateIST%
	utcTime := backupString
return

; Close ToolTip
TOOLTIP:
ToolTip,
SetTimer, TOOLTIP, Off
Return

; Get state of CapsLock
CapsLock_State_Toggle:
	If GetKeyState("CapsLock", "T")
	{
		state = Off
	}
	Else
	{
		state = On
	}
	CapsLock_State_Toggle(state)
Return

; Set the state of CapsLock
CapsLock_State_Toggle(State)
{
	SetCapsLockState, %State%
	ToolTip, CapsLock %State%
	SetTimer, TOOLTIP, On
}

; https://gist.github.com/davebrny/8bdbef225aedf6478c2cb6414f4b9bce
; Type what the anchor word is and swap the text on the opposite sides
text_swap(string)
{
	loop, % strLen(string) / 1.6
		div .= "- " ; Make divider between old string and new string
	mouseGetPos, mx, my
	swapped := string
	toolTip, % "swap at: """ . this . """`n`n" string "`n" div "`n" swapped, mx, my + 50 ; Display old string
	; Loop ToolTip until endkey is selected
	loop,
	{
		input, new_input, L1, {enter}{esc}{backspace}
		endkey := strReplace(errorLevel, "EndKey:", "")
		if endkey contains enter, escape
			break
		if (endkey = "backspace")
			stringTrimRight, this, this, 1
		if inStr(string, new_input)
			this .= new_input
		swapped := swap(string, this)
		tooltip, % "swap at: """ . this . """`n`n" string "`n" div "`n" swapped, mx, my + 50
	}
	tooltip, ; clear
	; Save swapped text
	if (this != "") and (endkey = "enter")
	{
		string := swapped
		sleep 300
	}
	this := ""
	div := ""
	return string
}

; Logic to swap text in string using the anchor string at_this
swap(string, at_this) {
	stringGetPos, pos, string, % at_this
	stringMid, left, string, pos, , L
	stringGetPos, pos, string, % at_this
	stringMid, right, string, pos + strLen(at_this) + 1
	stringRight, left_space, left, % strLen(left) - strLen(rTrim(left))
	stringLeft, right_space, right, % strLen(right) - strLen(lTrim(right))
	return lTrim(right) . left_space . at_this . right_space . rTrim(left)
}

; https://www.autohotkey.com/board/topic/73844-tabs-to-spaces-which-preserves-alignment/
; Convert Tabs to Spaces
TabsToSpaces(string, outEOL="`r`n", EOL="`n", Omit="`r"){
	Loop Parse, string, %EOL%, %Omit%
	{
		index := 0
		Loop Parse, A_LoopField
		{
			index++
			If (A_LoopField = A_Tab){
				Loop % 4-Mod(index, 4)
					r .= " "
				index := -1
			}
			else r .= A_LoopField
		}
	r .= outEOL
}
StringTrimRight, r, r, % StrLen(outEOL)
return r
}

; Convert Spaces to Tabs
SpacesToTabs(string, outEOL="`r`n", EOL="`n", Omit="`r"){
	Loop Parse, string, %EOL%, %Omit%
	{
		index := 0
		spaceCount := 0
		Loop Parse, A_LoopField
		{
			index++
			If (A_LoopField = A_Space){
				spaceCount++
				If (spaceCount == 4){
					r .= "	"
					; index := -1
					spaceCount := 0
				}
			}
			else r .= A_LoopField
		}
	r .= outEOL
}
StringTrimRight, r, r, % StrLen(outEOL)
return r
}

; https://www.autohotkey.com/boards/viewtopic.php?t=71645
; Return location of active File Explorer directory
ExplorerPath(_hwnd)
{
	for w in ComObjCreate("Shell.Application").Windows
		if (w.hwnd = _hwnd)
		{
			path := StrReplace(w.LocationURL, "file:///")
			replace := StrReplace(path, "%20", " ")
			replace := StrReplace(replace, "/", "\")
			return replace
		}
}

; https://www.computoredge.com/AutoHotkey/Downloads/QuickLinksTimeDateSubMenuSwitch.ahk
; Creates string of formatted dates
DateFormats(Date)
{
	FormatTime, OutputVar , %Date%, h:mm tt EST
	List := OutputVar
	FormatTime, OutputVar , %Date%, HH:mm EST
	List := List . "~" . OutputVar
	FormatTime, OutputVar , %Date%, ShortDate
	List := List . "~" . OutputVar
	FormatTime, OutputVar , %Date%, MMM. d, yyyy
	List := List . "~" . OutputVar
	FormatTime, OutputVar , %Date%, MMMM d, yyyy
	List := List . "~" . OutputVar
	FormatTime, OutputVar , %Date%, LongDate
	List := List . "~" . OutputVar
	FormatTime, OutputVar, %Date%, h:mm tt EST, dddd, MMMM d, yyyy
	List := List . "~" . OutputVar
	FormatTime, OutputVar, %Date%, dddd, MMMM d, yyyy @ hh:mm:ss tt EST
	List := List . "~" . OutputVar
	FormatTime, OutputVar, %Date%, ddd_MM-dd-yyyy_hh-mmtt_EST
	List := List . "~" . OutputVar
	Return List
}

; https://www.computoredge.com/AutoHotkey/Downloads/QuickLinksTimeDateSubMenuSwitch.ahk
; Creates DateTime Submenu
TextMenuDate(TextOptions, Menu, Action)
{
	StringSplit, MenuItems, TextOptions , ~
	Loop %MenuItems0%
	{
		Item := MenuItems%A_Index%
		Menu, %Menu%, add, %Item%, %Action%
		Switch
		{
		Case (InStr(Item,":") and InStr(Item,"`,")):
			Menu, TimeDate, Icon, %Item%, %A_Windir%\System32\timedate.cpl
		Case (InStr(Item,":")):
			Menu, TimeDate, Icon, %Item%, %A_Windir%\System32\shell32.dll, 240
		Default:
			Menu, TimeDate, Icon, %Item%, %A_Windir%\System32\ieframe.dll, 46
		}
	}
}

; https://www.computoredge.com/AutoHotkey/Downloads/QuickLinksTimeDateSubMenuSwitch.ahk
; DateTime submenu action
DateAction:
	SendInput %A_ThisMenuItem%{Raw}%A_EndChar%
Return

; https://www.autohotkey.com/boards/viewtopic.php?t=95931
; Return time object from API
time(area) {
	WinHttp := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	WinHttp.Open("GET", "https://worldtimeapi.org/api/timezone/" . area, false), WinHttp.Send()
	timeObject := JsonToAHK(WinHttp.ResponseText)
	Return timeObject
}

; https://www.autohotkey.com/boards/viewtopic.php?t=67583
; Convert JSON to AHK object
JsonToAHK(json, rec := false) {
	static doc := ComObjCreate("htmlfile")
		, __ := doc.write("<meta http-equiv=""X-UA-Compatible"" content=""IE=9"">")
		, JS := doc.parentWindow
	if !rec
		obj := %A_ThisFunc%(JS.eval("(" . json . ")"), true)
	else if !IsObject(json)
		obj := json
	else if JS.Object.prototype.toString.call(json) == "[object Array]" {
		obj := []
		Loop % json.length
			obj.Push( %A_ThisFunc%(json[A_Index - 1], true) )
	}
	else {
		obj := {}
		keys := JS.Object.keys(json)
		Loop % keys.length {
			k := keys[A_Index - 1]
			obj[k] := %A_ThisFunc%(json[k], true)
		}
	}
	Return obj
}

; Exit app
QUIT:
ExitApp
