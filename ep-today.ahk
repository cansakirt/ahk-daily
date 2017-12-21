#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


;Changelog 12/04/2017
;Created changelog
; 12/04 Merged two different versions from my desk and nina's office. 
; 12/04 I shall use the email draft to keep a live copy for further changes and syncing.
; 12/04 Added IfWinActive with ahk_class
; 12/05 Tried and commented out cycling thru shortcuts for firstclass colors
; 12/05 Added banner login
; 12/12 Added shift+alt+L link & replaced %clipboard% with dedicated variables
; 12/20 -CS capitalizes and does not convert to epdocx
; 12/20 Added a check if Clipboard starts with @ before renaming, else warning & Keeps DCEID in clipboard after renaming folders

; ------------------- BEGIN sis ----------------
#IfWinActive, ahk_class SunAwtFrame
^e::
SendInput TERM{tab}ACT_CODE{tab}{ShiftDown}{tab}{shiftup}{CtrlUp}
Return
^!F9:: ; ctrl+alt+f9 for login input
SendInput username{TAB}password{enter}
Return
#IfWinActive
; ------------------- END sis ------------------
; ------------------- BEGIN Windows Explorer ----------------
#IfWinActive, ahk_class CabinetWClass
^.:: ; COPY ID FIRST ; rename email folder to lastname_firstname  ID  date
dceid := Clipboard
stringleft, startwithAt, dceid, 1
if startwithAt = @
{
;MsgBox, it's gonna work. id: %dceid% - @: %startwithAt%
SendInput {f2}{end}{CtrlDown}{Left 5}{CtrlUp}{ShiftDown}{home}{CtrlDown}{right 2}{CtrlUp}{Shiftup}{Space 4}^v{Space 4}{CtrlDown}{Left}{CtrlUp}{Left 5}{ShiftDown}{Ctrldown}{left}{CtrlUp}{Shiftup}^x{home}{Delete 2}^v{enter}
Clipboard := dceid
dceid =
}
else 
MsgBox, %dceid% `r`n is not a DCE ID!	;completely optional
Return
^d::  ; rename selected file into decision.txt
SendInput {F2}Decision{ENTER}
Return
+^d::
SendInput {F2}DECISION{ENTER}
Return
#IfWinActive
; ------------------- END Windows Explorer ------------------


; ------------------- BEGIN FirstClass ----------------------
#IfWinActive, ahk_class SAWindow
!b:: ; FirstClass blue color
SendInput {alt}rc{up 3}{enter}
Return
!o:: ; FirstClass orange color
SendInput {alt}rc{down 4}{enter}
Return
!s:: ; some signature in FirstClass
SendInput {alt}eic
Return
!l:: ; alt+L to make link to EP page
SendInput {alt}el
epwebpage = https://example.site
SendInput %epwebpage%
Return
+!l:: ; shift+alt+L to make link to EP page
SendInput {alt}el
wowsuch = https://other.site/
SendInput %wowsuch%
Return

^+v::
Clipboard=%Clipboard%   ; will remove formatting
Sleep, 100   ; wait for Clipboard to update
Send ^v
Return


; `:: ; cycle thru colors.
; i++
; if (i>3)
    ; i=0
; else if (i=1)
   ; SendInput {alt}rc{up 3}{enter}
; blue
; else if (i=2)
     ; SendInput {alt}rc{down 4}{enter}
; orange
; else if (i=3)
; SendInput	{alt}rc{enter} ; black
; return
#IfWinActive
; ------------------- END FirstClass ------------------------

; ------------------- BEGIN Notepad++ ------------------------
#IfWinActive ahk_exe notepad++.exe
Pause:: ;I did this note. -initials 11/09
today = %a_now%
FormatTime, today, %today%, MM/dd
SendInput I did this, this is a note. -XX %today%
return
!Pause:: ; looks fine
today = %a_now%
FormatTime, today, %today%, MM/dd
SendInput Looks fine to me. -XX %today%
Return
::-xx::
SendInput -XX ;if I sign off with initials don't do the thing right below.
Return
::xx::
today = %a_now%
today1 = %a_now%
FormatTime, today, %today%, MM/dd
FormatTime, today1, %today1%, MM/dd/yyyy
SendInput XX{TAB}%today1%{enter}ACT_CODE{TAB}%today%
Return
#IfWinActive
; ------------------- END Notepad++ ------------------------


;::++::
;Send, %A_MM%/%A_DD%/%A_YYYY%
;Return

NumpadMult::  ; 11/09 today
today = %a_now%
FormatTime, today, %today%, MM/dd
SendInput %today%   
return

NumpadSub::  ; 09-Nov-2017 today
today = %a_now%
FormatTime, today, %today%, dd-MMM-yyyy
SendInput %today%   
return

#IfWinActive ahk_exe chrome.exe
^e::  ; search in example.site on google
SendInput {space}site:example.site{space}
Return
#IfWinActive

^r::
Reload
Sleep, 500
msgbox,,, DIDN'T WORK
return
