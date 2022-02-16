#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
; Source/Inspiration https://www.autohotkey.com/board/topic/4310-capshift-slow-down-and-extend-the-caps-lock-key/

/*
***************************************** 
TODO:
Insert time menu to insert xx:xx time for diff timezones
Storage Usage
Indicate what processes/programs are using this file (handles)

*******************************************
*/

About =
(LTrim0
	Slows down and extends the capslock key. Hold for 0.05 sec to toggle capslock on or off. Hold for 0.3 sec to show a menu that converts selected text to UPPER CASE, lower case, Title Case, iNVERT cASE, etc. Also can trigger via Right Click + Middle Click.
	)

; Declare variables
SetTimer,TOOLTIP,1500
SetTimer,TOOLTIP,Off
TimeCapsToggle = 5
TimeOut = 30
global string2 := ""

; Define hotstring, trigger via hold CapsLock for 1.5s or hold Right Click + Middle click
CapsLock::
~RButton & MButton::
	OldClipboard:= ClipboardAll ;Save existing clipboard.
	counter = 0
	Progress, ZH16 ZX0 ZY0 B R0-%TimeOut%
	Loop, %TimeOut%
	{
		Sleep, 10
		counter += 1
		Progress, %counter% ;, SubText, MainText, WinTitle, FontName
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

	if (string2){
		clipboard := string2
	} else {
		Clipboard := OldClipboard
	}
Return

; Build CapsLock menu
MENU:
	Winget, Active_Window, ID, A
	Send, ^c
	ClipWait, 1
	Menu, convert, Add
	Menu, convert, Delete
	Menu, misc, Add, &About, ABOUT
	Menu, misc, Add, &Quit, QUIT
	Menu, convert, Add, CAPshift, :misc
	Menu, convert, Add, 	
	Menu, convert, Add, &CapsLock Toggle, CapsLock_State_Toggle
	If (GetKeyState("CapsLock", "T") = True)
		Menu,Convert,Check, &CapsLock Toggle
	Else
		Menu,Convert,uncheck, &CapsLock Toggle	
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
	Menu, Browser Search..., Add, &Google, MENU_ACTION
	Menu, Browser Search..., Add, &Thesaurus, MENU_ACTION
	Menu, Browser Search..., Add, &Wikipedia, MENU_ACTION
	Menu, Browser Search..., Add, &Define, MENU_ACTION
	Menu, Browser Search..., Add, &Open Page..., MENU_ACTION
	Menu, convert, Add, &Browser Search..., :Browser Search...
	Menu, convert, Add, 
	Menu, Explorer..., Add, &Open Folder..., MENU_ACTION	
	Menu, Explorer..., Add, &Open Parent Folder..., MENU_ACTION	
	Menu, Explorer..., Add, &Copy File Names and Details from Folder to Clipboard..., MENU_ACTION	
	Menu, Explorer..., Add, &Add File to Subfolder..., MENU_ACTION
	Menu, Explorer..., Add, &Pull File Out to Parent Folder..., MENU_ACTION
	Menu, Explorer..., Add, &Clipboard to File...., MENU_ACTION	
	Menu, Explorer..., Add, &File to Clipboard..., MENU_ACTION
	Menu, convert, Add, &Explorer..., :Explorer...
	Menu, convert, Add, 
	Menu, Scripts..., Add, &CCleaner, MENU_ACTION
	Menu, Scripts..., Add, &RemoveResourceGroups.ps1, MENU_ACTION
	Menu, Scripts..., Add, &createNewVM.ps1, MENU_ACTION
	Menu, Scripts..., Add, &{Update NSGS}.ps1, MENU_ACTION		
	Menu, Scripts..., Add, &getDiskSpace.ps1, MENU_ACTION				
	Menu, convert, Add, &Scripts..., :Scripts...
	Menu, convert, Add, 
	Menu, TimeDate, Add, Time/Date, MENU_ACTION
	Menu, TimeDate, DeleteAll
	List := DateFormats(A_Now)
	TextMenuDate(List,"TimeDate","DateAction")
	Menu, convert, Add, &Insert Time/Date, :TimeDate
	Menu, convert, Add, &Insert Specified Time, MENU_ACTION
	Menu,convert, Default, &CapsLock Toggle	
	Menu, convert, Show
Return

; Pass copied string to specified menu action, then paste if required
MENU_ACTION:
	AutoTrim, Off
	string = %clipboard%
	clipboard := Menu_Action(A_ThisMenuItem, string)
	WinActivate, ahk_id %Active_Window%
	If (!string2){
		Send, ^v
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
		StringCaseSense,On
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
		StringCaseSense,On
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
	; Convert phrase separated_by_or-to lowerCamelCase
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
		Run, http://www.google.com/search?q=%string%
	; Looks the highlighted word up in a thesaurus
	Else If ThisMenuItem =&Thesaurus
		Run, http://www.thesaurus.com/browse/%string%
	; Looks the highlighted word up in Wiki
	Else If ThisMenuItem =&Wikipedia
		Run, https://en.wikipedia.org/wiki/%string% 
	; Defines the highlighted word
	Else If ThisMenuItem =&Define
		Run, http://www.google.com/search?q=define+%string%
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
		SplitPath, string,, dir
		Run explorer.exe /select`,%dir%
	}
	; If the copied text is a valid site, open it in web browser
	Else If ThisMenuItem =&Open Page...
		Run, %string% 
	; Run old portable Ccleaner in auto mode
	Else If ThisMenuItem =&CCleaner
		Run, "D:\OneDrive\APPS\CCleaner\CCleaner.exe" /auto
	Else If (ThisMenuItem ="&Copy File Names and Details from Folder to Clipboard...") {
		; https://lexikos.github.io/v2/docs/FAQ.htm#output
		SplitPath, string, , dir
		string2Temp := dir . "\string2.txt"
		RunWait %comspec% ' "/c dir "%dir%" /a /o /q /t > "%string2Temp%" "', , hide
		string2File := FileOpen(string2Temp, "rw")
		string2 := string2File.Read()
		string2File.Close()
		FileDelete, %string2Temp%
	}
	Else If (ThisMenuItem ="&File to Clipboard..."){
		file := FileOpen(string, "r")	
		string2 := File.Read()
		file.Close()
	}
	; Run getDiskSpace.ps1 script
	Else If ThisMenuItem =&getDiskSpace.ps1
		Run, powershell -NoExit -File "D:\Dropbox\code\getDiskSpace.ps1"
	; Run RemoveResourceGroups.ps1 script
	Else If ThisMenuItem =&RemoveResourceGroups.ps1
		Run, powershell -NoExit -File "C:\Users\rymccall\OneDrive - Microsoft\PowerShell\byronbayer\Powershell\Remove-ResourceGroupsAsyncWithPattern.ps1"
	; Run createNewVM.ps1 script
	Else If ThisMenuItem =&createNewVM.ps1
		Run, C:\Users\rymccall\GitHub\Azure-PowerShell---Create-New-VM\createNewVM.ps1
	; Run {Update NSGS}.ps1 script
	Else If ThisMenuItem =&{Update NSGS}.ps1
		Run, powershell -NoExit -File "C:\Users\rymccall\OneDrive - Microsoft\PowerShell\Update your NSGs with a rule scoped to your IP address.ps1"
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
	Else If ThisMenuItem =&Spaces to Tabs
	{		
		; string := StrReplace(string, "    ", "	", , Limit := -1)
		string := RegExReplace(string, "[\s]{4}", A_Tab)
		; string := SpacesToTabs(string)
	}
	; Create a new subfolder, add file to it
	Else If ThisMenuItem =&Add File to Subfolder...
	{
		SplitPath, string, , dir, , nameNoExt
		FileCreateDir, %nameNoExt%
		FileMove, %string%, %nameNoExt%
	}

	; Take file out of current folder and add to parent folder
	Else If ThisMenuItem =&Pull File Out to Parent Folder...
	{
		SplitPath, string,, dir
		SplitPath, dir,, dir2
		FileMove, %string%, %dir2%
	}

	; Copy clipboard contents to a text file in a folder currently open in Windows Explorer
	Else If ThisMenuItem =&Clipboard to File....
	{
		WinGet, WinID, ID, ahk_class CabinetWClass
		CurrentPath := ExplorerPath(WinID)
		newPath := CurrentPath "\clipboardToFile.txt"
		; FileAppend, string`n, %CurrentPath%%newPath%
		file := FileOpen(newPath, "rw")
		file.Write(string "`n")
		file.Close()
	}
	; Convert time and paste string in different time zones
	Else If ThisMenuItem =&Insert Specified Time
	{
		string := timeConversion(string)
	}
Return string
}

; N/A
EMPTY:
Return

; Close ToolTip
TOOLTIP:
	ToolTip,
	SetTimer,TOOLTIP,Off
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

; Display About message
ABOUT:
	MsgBox, 0, CAPshift, %About%
Return

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

; Convert the time
timeConversion(string)
{
	If string is not time
	{
		string := A_NowUTC
	} else {
		string := DateParse(string)
	}
	backupString := string
	estObject := time("America/New_York")
	cstObject := time("America/Chicago")
	pstObject := time("PST8PDT")
	estOffset := estObject.utc_offset
	cstOffset := cstObject.utc_offset
	pstOffset := pstObject.utc_offset
	istOffset := "05:30"

	string += estOffset, hours
	FormatTime, OutputVar, %string%, MMM. d @ h:mmtt
	stringTime := OutputVar . " EST ["
	string := backupString

	string += cstOffset, hours
	FormatTime, OutputVar, %string%, MMM. d @ h:mmtt
	stringTime := stringTime . OutputVar . " CST | "
	string := backupString

	string += pstOffset, hours
	FormatTime, OutputVar, %string%, MMM. d @ h:mmtt
	stringTime := stringTime . OutputVar . " PST | "
	string := backupString

	string += istOffset, hours
	FormatTime, OutputVar, %string%, MMM. d @ h:mmtt
	stringTime := stringTime . OutputVar . " IST]"
	string := backupString

	; stringTime := string
	; stringTime += "`n"
	; stringTime += estOffset, hours
	; stringTime += "`n"

return stringTime
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
	FormatTime, OutputVar , %Date%, h:mm tt ;12 hour clock
	List := OutputVar
	FormatTime, OutputVar , %Date%, HH:mm ;24 hour clock
	List := List . "~" . OutputVar
	FormatTime, OutputVar , %Date%, ShortDate ; 11/5/2015
	List := List . "~" . OutputVar
	FormatTime, OutputVar , %Date%, MMM. d, yyyy
	List := List . "~" . OutputVar
	FormatTime, OutputVar , %Date%, MMMM d, yyyy
	List := List . "~" . OutputVar
	FormatTime, OutputVar , %Date%, LongDate
	List := List . "~" . OutputVar
	FormatTime, OutputVar, %Date%, h:mm tt, dddd, MMMM d, yyyy
	List := List . "~" . OutputVar
	FormatTime, OutputVar, %Date%, dddd, MMMM d, yyyy @ hh:mm:ss tt
	List := List . "~" . OutputVar
	FormatTime, OutputVar, %Date%, ddd_MM-dd-yyyy_hh-mmtt_EST
	List := List . "~" . OutputVar
Return List
}

; https://www.computoredge.com/AutoHotkey/Downloads/QuickLinksTimeDateSubMenuSwitch.ahk
; Creates DateTime Submenu
TextMenuDate(TextOptions,Menu,Action)
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
	WinHttp.Open("GET", "http://worldtimeapi.org/api/timezone/" area, false), WinHttp.Send()
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
return

; https://www.autohotkey.com/board/topic/18760-date-parser-convert-any-date-format-to-yyyymmddhh24miss
; Convert date string to date object
DateParse(str, americanOrder=0) {
	static monthNames := "(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*"
	, dayAndMonth := "(?:(\d{1,2}|" . monthNames . ")[\s\.\-\/,]+)?(\d{1,2}|" . monthNames . ")"
	If RegExMatch(str, "i)^\s*(?:(\d{4})([\s\-:\/])(\d{1,2})\2(\d{1,2}))?"
			. "(?:\s*[T\s](\d{1,2})([\s\-:\/])(\d{1,2})(?:\6(\d{1,2})\s*(?:(Z)|(\+|\-)?"
	. "(\d{1,2})\6(\d{1,2})(?:\6(\d{1,2}))?)?)?)?\s*$", i) ;ISO 8601 timestamps
	year := i1, month := i3, day := i4, t1 := i5, t2 := i7, t3 := i8
	Else If !RegExMatch(str, "^\W*(\d{1,2}+)(\d{2})\W*$", t){
		RegExMatch(str, "i)(\d{1,2})"					;hours
			. "\s*:\s*(\d{1,2})"				;minutes
			. "(?:\s*:\s*(\d{1,2}))?"			;seconds
		. "(?:\s*([ap]m))?", t)				;am/pm
		StringReplace, str, str, %t%
		If Regexmatch(str, "i)(\d{4})[\s\.\-\/,]+" . dayAndMonth, d) ;2004/22/03
			year := d1, month := d3, day := d2
		Else If Regexmatch(str, "i)" . dayAndMonth . "[\s\.\-\/,]+(\d{2,4})", d) ;22/03/2004 or 22/03/04
			year := d3, month := d2, day := d1
		If (RegExMatch(day, monthNames) or americanOrder and !RegExMatch(month, monthNames)) ;try to infer day/month order
			tmp := month, month := day, day := tmp
	}
	f = %A_FormatFloat%
	SetFormat, Float, 02.0
	d := (StrLen(year) == 2 ? "20" . year : (year ? year : A_YYYY))
	. ((month := month + 0 ? month : InStr(monthNames, SubStr(month, 1, 3)) // 4 ) > 0 ? month + 0.0 : A_MM)
	. ((day += 0.0) ? day : A_DD) 
	. t1 + (t1 == 12 ? t4 = "am" ? -12.0 : 0.0 : t4 = "pm" ? 12.0 : 0.0)
	. t2 + 0.0 . t3 + 0.0
	SetFormat, Float, %f%
return, d
}

; Exit app
QUIT:
ExitApp
