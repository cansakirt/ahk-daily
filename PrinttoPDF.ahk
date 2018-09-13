#useHook On
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent
file = %1%
extension := SubStr(file, -3)
; msgbox %extension%
if extension != .pdf
{
	msgbox This is not a pdf file.`nExiting... ;`n%file%`n%extension%
	exitapp
}
else
; msgbox, This is the file:`n%1%
run AcroRd32.exe /t "%1%" "Microsoft Print to PDF",,,OutputVarPID
winwait, Save Print Output As
; msgbox, %OutputVarPID%
sleep, 150

filename = %1%
; msgbox %filename%
length := strlen(filename)
; msgbox %length%
filename := SubStr(filename, 1, length-4)
filename = %filename%_COPY.pdf
; msgbox %filename%
Send {Text}%filename% ;, Save Print Output As
Send {enter} ;, Save Print Output As
winwait, Adobe Acrobat Reader DC
process, close, %OutputVarPID%
; msgbox, %OutputVarPID%
exitapp
