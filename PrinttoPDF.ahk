#useHook On
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent
file = %1%
; msgbox % file
if 0 > 1
{
	msgbox, %0% files will be printed.
	Loop %0%  ; For each parameter (or file dropped onto a script):
	{
		GivenPath := %A_Index%  ; Fetch the contents of the variable whose name is contained in A_Index.
		MsgBox Now printing:`n%GivenPath%
		file := %A_Index%
		0 -= 1
		genesis(file)
	}
}
if 0 = 1
{
	; msgbox printing one file.
	genesis(file)
}
; msgbox end of run.
exitapp

genesis(file)
{
	extension := SubStr(file, -3)
	; msgbox %extension%
	if extension != .pdf
	{
		msgbox Please drag and drop a PDF file.`nExiting... ;`n%file%`n%extension%
		exitapp
	}
	else
	; msgbox, This is the file:`n%file%
	run AcroRd32.exe /t "%file%" "Microsoft Print to PDF",,,OutputVarPID
	winwait, Save Print Output As
	; msgbox, %OutputVarPID%
	sleep, 150

	filename = %file%
	; msgbox %filename%
	length := strlen(filename)
	; msgbox %length%
	filename := SubStr(filename, 1, length-4)
	filename = %filename%_COPY.pdf
	; msgbox %filename%
	winactivate, Save Print Output As
	Send {Text}%filename% ;, Save Print Output As
	Send {enter} ;, Save Print Output As
	winwait, Adobe Acrobat Reader DC
	process, close, %OutputVarPID%
	; msgbox, %OutputVarPID%
}
