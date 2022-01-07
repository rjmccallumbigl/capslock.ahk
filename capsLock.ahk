#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
; Source https://www.autohotkey.com/board/topic/4310-capshift-slow-down-and-extend-the-caps-lock-key/page-2

About =
(LTrim0
	Slows down and extends the capslock key. Hold for 0.05 sec to toggle capslock on or off. Hold for 0.3 sec to show a menu that converts selected text to UPPER CASE, lower case, Title Case, iNVERT cASE, etc. Also can trigger via Right Click + Middle Click.
	)

SetTimer,TOOLTIP,1500
SetTimer,TOOLTIP,Off
TimeCapsToggle =5
TimeOut =30
CapsLock::
~RButton & MButton::
	OldClipboard:= ClipboardAll ;Save existing clipboard.
	counter=0
	Progress, ZH16 ZX0 ZY0 B R0-%TimeOut%
	Loop, %TimeOut%
	{
		Sleep,10
		counter+=1
		Progress, %counter% ;, SubText, MainText, WinTitle, FontName
		If (counter = TimeCapsToggle)
			Progress, ZH16 ZX0 ZY0 B R0-%TimeOut% CBFF0000
		if (A_ThisHotkey = "CapsLock")
		{
			GetKeyState,state,CapsLock,P
			If state=U
				Break
		}

		Else If (A_ThisHotkey = "MButton")
		{
			GetKeyState,state,MButton,P
			If state=U
				Break
		}
	}
	Progress, Off
	If counter=%TimeOut%
		Gosub,MENU
	Else If (counter>TimeCapsToggle)
	{		
		Gosub, CapsLock_State_Toggle
	}

	Clipboard:= OldClipboard
Return

MENU:
	Winget, Active_Window, ID, A
	Send,^c
	ClipWait,1
	Menu,convert,Add
	Menu,convert,Delete
	Menu, misc ,Add,&About,ABOUT
	Menu, misc ,Add,&Quit,QUIT
	Menu,convert,Add,CAPshift, :misc
	Menu,convert,Add,	
	Menu,Convert,Add,&CapsLock Toggle,CapsLock_State_Toggle	
	If (GetKeyState("CapsLock", "T") = True )
		Menu,Convert,Check,&CapsLock Toggle
	Else
		Menu,Convert,uncheck,&CapsLock Toggle	
	Menu,convert,Add,
	Menu,convert,Add,&UPPER CASE,MENU_ACTION
	Menu,convert,Add,&lower case,MENU_ACTION
	Menu,convert,Add,
	Menu,convert,Add,&Title Case,MENU_ACTION
	Menu,convert,Add,&Sentence case,MENU_ACTION
	Menu,convert,Add,&iNVERT cASE,MENU_ACTION
	Menu,convert,Add,&SpOnGeBoB,MENU_ACTION
	; Menu,convert,Add,&HaLf CaPs CaSe,MENU_ACTION
	Menu,convert,Add,
	Menu,convert,Add,Remove_&under_scores,MENU_ACTION
	Menu,convert,Add,Remove.&full.stops,MENU_ACTION
	Menu,convert,Add,
	Menu,convert,Add,&Google,MENU_ACTION
	Menu,convert,Add,&Thesaurus,MENU_ACTION
	Menu,convert,Add,&Wikipedia,MENU_ACTION
	Menu,convert,Add,&Define,MENU_ACTION
	Menu,convert,Add,
	Menu,convert,Add,&Open Folder...,MENU_ACTION
	Menu,convert,Add,&Open Page...,MENU_ACTION
	Menu,convert,Add,&Open Parent Folder...,MENU_ACTION
	Menu,convert,Add,
	Menu,convert,Add,&Clean,MENU_ACTION
	Menu,convert,Add,&RemoveResourceGroups.ps1,MENU_ACTION
	Menu,convert,Add,&createNewVM.ps1,MENU_ACTION
	Menu,convert,Add,&{Update NSGS}.ps1,MENU_ACTION	
	Menu,convert,Default,&CapsLock Toggle
	Menu,convert,Show
Return

MENU_ACTION:
	AutoTrim,Off
	string=%clipboard%
	clipboard:=Menu_Action(A_ThisMenuItem, string)
	WinActivate, ahk_id %Active_Window%
	Send,^v
	ToolTip,Selection converted to %A_ThisMenuItem%
	SetTimer,TOOLTIP,On
Return

Menu_Action(ThisMenuItem, string)
{
	If ThisMenuItem =&UPPER CASE
		StringUpper,string,string

	Else If ThisMenuItem =&lower case
		StringLower,string,string

	Else If ThisMenuItem =&Title Case
		StringLower,string,string,T
	Else If ThisMenuItem =&SpOnGeBoB
	{
		StringCaseSense,On
		newString := ""
		splitString:= StrSplit(string)
		for index, element in splitString
		{
			If (Mod(index,2) = 0)
			{
				StringUpper, newChar, element
				newString .= newChar
			} Else{
				StringLower, newChar, element
				newString .= newChar
			}
		}
		string := newString
		StringCaseSense,Off
	}	

	Else If ThisMenuItem =&iNVERT cASE
	{
		StringCaseSense,On
		lower=abcdefghijklmnopqrstuvwxyz
		upper=ABCDEFGHIJKLMNOPQRSTUVWXYZ
		StringLen,length,string
		Loop,%length%
		{
			StringLeft,char,string,1
			StringGetPos,pos,lower,%char%
			pos+=1
			If pos<>0
				StringMid,char,upper,%pos%,1
			Else
			{
				StringGetPos,pos,upper,%char%
				pos+=1
				If pos<>0
					StringMid,char,lower,%pos%,1
			}
			StringTrimLeft,string,string,1
			string.=char
		}
		StringCaseSense,Off
	}

	; HaLf CaPs CaSe, only works with one word
	Else If ThisMenuItem =&HaLf CaPs CaSe		
	{
		StringLower, string, string, T
		StringCaseSense,On
		Loop, Parse, string

		lower= abcdefghijklmnopqrstuvwxyz
		upper= ABCDEFGHIJKLMNOPQRSTUVWXYZ
		StringLen,length,string
		Loop,%length%
		{
			StringLeft,char,string,1
			StringGetPos,pos,lower,%char%
			pos+=1
			If (Mod(counter, 2) = 0)
				StringMid,char,upper,%pos%,1
			Else
			{
				StringGetPos,pos,upper,%char%
				pos+=1
				If (Mod(counter, 2) = 0) 
					; If (%char% != " ")
				StringMid,char,lower,%pos%,1
				; }
			}
			StringTrimLeft,string,string,1
			string.=char
			counter++
		}
		StringCaseSense,Off
	}

	Else If ThisMenuItem =&Sentence case
	{
		StringCaseSense,On
		lower=abcdefghijklmnopqrstuvwxyz
		upper=ABCDEFGHIJKLMNOPQRSTUVWXYZ
		dot=1
		StringLen,length,string
		Loop,%length%
		{
			StringLeft,char,string,1
			if (char==".")
			{
				dot=1
			}
			else
			{
				if (dot=1)
				{
					StringGetPos,pos,lower,%char%
					pos+=1
					If pos<>0
						StringMid,char,upper,%pos%,1
				}
				else
				{
					StringGetPos,pos,upper,%char%
					pos+=1
					If pos<>0
						StringMid,char,lower,%pos%,1
				}

				if (char<> " ")
					dot=0
			}
			StringTrimLeft,string,string,1
			string.=char
		}
		StringCaseSense,Off
	}
	Else If ThisMenuItem =Remove_&under_scores
		StringReplace, string, string,_,%A_Space%, All
	Else If ThisMenuItem =Remove.&full.stops
		StringReplace, string, string,.,%A_Space%, All
	Else If ThisMenuItem =&Google
		Run, http://www.google.com/search?q=%string%
	Else If ThisMenuItem =&Thesaurus
		Run, http://www.thesaurus.com/browse/%string%
	Else If ThisMenuItem =&Wikipedia
		Run, https://en.wikipedia.org/wiki/%string% 
	Else If ThisMenuItem =&Define
		Run, http://www.google.com/search?q=define+%string%
	Else If ThisMenuItem =&Open Folder...
	{

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
	Else If ThisMenuItem =&Open Parent Folder...
	{
		SplitPath, string,, dir
		Run explorer.exe /select`,%dir%
	}
	Else If ThisMenuItem =&Open Page...
		Run, %string% 
	Else If ThisMenuItem =&Clean
		Run, "D:\OneDrive\APPS\CCleaner\CCleaner.exe" /auto
	Else If ThisMenuItem =&RemoveResourceGroups.ps1
		Run, powershell -NoExit -File "C:\Users\rymccall\OneDrive - Microsoft\PowerShell\byronbayer\Powershell\Remove-ResourceGroupsAsyncWithPattern.ps1"
	Else If ThisMenuItem =&createNewVM.ps1
		Run, C:\Users\rymccall\GitHub\Azure-PowerShell---Create-New-VM\createNewVM.ps1
	Else If ThisMenuItem =&{Update NSGS}.ps1
		Run, powershell -NoExit -File "C:\Users\rymccall\OneDrive - Microsoft\PowerShell\Update your NSGs with a rule scoped to your IP address.ps1"

Return string
}

EMPTY:
Return

TOOLTIP:
	ToolTip,
	SetTimer,TOOLTIP,Off
Return

CapsLock_State_Toggle:
	If GetKeyState("CapsLock","T")
	{
		state=Off
		; Menu,Convert,Uncheck,&CapsLock Toggle
	}		
	Else
	{
		state=On
		; Menu,Convert,Check,&CapsLock Toggle
	}
	CapsLock_State_Toggle(state)

Return

CapsLock_State_Toggle(State)
{
	SetCapsLockState,%State%	
	ToolTip,CapsLock %State%
	SetTimer,TOOLTIP,On
}

ABOUT:
	MsgBox,0,CAPshift,%About%
Return

QUIT:
ExitApp
