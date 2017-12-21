;Changelog 11/27/2017
;added IfWinActive with ahk_class
;added down arrow DOB color match function
;fixed not in sis check by adding sleep&send{f10}
;I shall use the email draft to keep a live copy for further changes and syncing.
; 12/19 added gosub for some things. trying out the NIB double check automation 
;


#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

^r::
Reload
Sleep, 500
msgbox,,, DIDN'T WORK
return


#IfWinActive, ahk_class SunAwtFrame
{
F12::	; not in sis Check
PixelGetColor, color, 925, 405 ; DOB match
; PixelGetColor, color, 260, 405 ; L Name match
; PixelGetColor, color, 700, 405 ; End of list msgbox
if color = 0x8CFFFF
{
MsgBox Birthday match! Color: %color%.
sleep, 300
send {tab}{tab}{space} ;Check the box
sleep, 300
send ^{PGUP} ;CTRL+PGUP
sleep, 200 
Send {tab} ;focus on the checkbox
Sleep, 30
Send {SPACE} ; UNcheck the box
Sleep, 30
Send {F10}
Send ^+{PGDN} ;CTRL+SHIFT+PGDN for Next set of Records
sleep, 100
send ^{PGDN} ;CTRL+PGDN
}
Else
gosub, nextone	;does what F12 does.
return

Down::  ; Down arrow hotkey.
gosub, labeldown
Return

LabelDown:
;MouseGetPos, MouseX, MouseY
PixelGetColor, color, 925, 405 ; DOB match
; PixelGetColor, color, 260, 405 ; L Name match
; PixelGetColor, color, 700, 405 ; End of list msgbox
if color = 0x8CFFFF
{
MsgBox Birthday match! Color: %color%.
}
Else Send {down}
return
}


#IfWinActive, ahk_class SunAwtFrame
{
; If found a match click NumpadAdd, the box checked tabtabspace 
NumpadAdd:: ;Brings up the next set of records
sleep, 300
send {tab}{tab}{space} ;Check the box
sleep, 500
send {F10} ;F10 to save and wait 700ms before moving on.
Sleep, 700
send {enter} ;Confirm F10 save msgbox
Sleep, 200
send ^{PGUP} ;CTRL+PGUP
sleep, 200 
Send ^+{PGDN} ;CTRL+SHIFT+PGDN for Next set of Records
sleep, 100
send ^{PGDN} ;CTRL+PGDN
return
}
#IfWinActive, ahk_class SunAwtFrame
{
; If no match found, click NumpadSub
NumpadSub:: ;checks "not in sis" box and brings up the next set of records
sleep, 300
send ^{PGUP} ;CTRL+PGUP
sleep, 200 
Send {tab} ;focus on the checkbox
Sleep, 30
Send {SPACE} ;check the box
Sleep, 30
Send {F10}
Send ^+{PGDN} ;CTRL+SHIFT+PGDN for Next set of Records
sleep, 100
send ^{PGDN} ;CTRL+PGDN
return

}
#IfWinActive, ahk_class SunAwtFrame
{
Numpad0::  ; 
today = %a_now%
today += -7, days
FormatTime, today, %today%, dd-MMM-yyyy 
SendInput %today%   
return
}


; EXPERIMENTAL

F1::	; scroll down the list see the matches, don't make any changes.
gosub, nextone
Return

nextone:
Sleep, 200
send ^{PGUP} ;CTRL+PGUP
sleep, 200 
Send ^+{PGDN} ;CTRL+SHIFT+PGDN for Next set of Records
sleep, 100
send ^{PGDN} ;CTRL+PGDN
Return



