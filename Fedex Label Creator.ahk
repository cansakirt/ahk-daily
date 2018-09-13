#useHook On
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance force
SetControlDelay, -1

{	; Initialization...
gosub, theLoading
; SetTimer, theloading, 1000
; SetTimer, theloading, off

; list := "can|berk|lan|ops" ; demo list for tests
Menu Tray, Icon, %A_ScriptDir%\favicon.ico
; MAIN GUI
gui +HWND_2hwnd
Gui Add, Button, x16 y8 w80 h35 gLoad -tabstop, Load File
gui, font, s12
Gui Add, DropDownList, hwndhwndvar x17 y65 w240 vList gUpdate, % list
Gui Add, Button, x264 y8 w80 h55 gPrev, Previous
Gui Add, Button, x264 y62 w80 h55 gNext, Next
Gui Add, Button, x264 y120 w80 h36 gInfo, More ⮃
gui, font
Gui Add, Button, x16 y96 w80 h55 gFedex, FedEx!
Gui Add, Text, x104 y4 w44 h21 +0x200, EDU ID:
Gui Add, Edit, x104 y21 w129 h21 vEDUidBox +hwnd_EduID

Gui Show, w360 h164 x1000 y200, FEDEX Label Creator v0.3

FilePath := A_Desktop "\fedexList.xlsx" ; example path
oWorkbook := ComObjGet(FilePath) ; access Workbook object
IDcolRange := oWorkbook.Sheets(1).Range("A1:A300")
gosub, load2
; dont forget to uncomment these 4 preceding lines.
Gui, Add,  Groupbox, x10 y190 w340 h292	+center , Mailing Information
; SECOND Gui
Gui, Add, Tab, -Background -tabstop x+200, Main|Out
gui, Tab, 1 
; Gui +HwndMyGuiHwnd
; gui, 2:New, +HWND_2hwnd -MinimizeBox, Student mailing information
Gui, Add, Text,  x16 y208 w80 h23 +0x200 vcol1, Country code
Gui, Add, Text,  xp yp+24 w80 h23 +0x200 vcol2, Contact name
Gui, Add, Text,  xp yp+24 wp hp +0x200 vcol3, Address 1
Gui, Add, Text,  xp yp+24 wp hp +0x200 vcol4, Address 2
Gui, Add, Text,  xp yp+24 wp hp +0x200 vcol5, Address 3
Gui, Add, Text,  xp yp+24 wp hp +0x200 vcol6, ZIP
Gui, Add, Text,  xp yp+24 wp hp +0x200 vcol7, City
Gui, Add, Text,  xp yp+24 wp hp +0x200 vcol8, Phone
Gui, Add, Text,  xp yp+24 wp hp +0x200 vcol9, Email
Gui, Add, Text,  xp yp+24 wp hp +0x200 vcol0, Fedexlink
gui, font, s9, Courier New
Gui, Add, Edit,  x96 y208 w35 h21 +uppercase ReadOnly vCCountry gCustomFedexCreate
Gui, Add, Edit,  x96 y232 w245 h21 ReadOnly vCName gCustomFedexCreate
Gui, Add, Edit,  xp yp+24 wp hp ReadOnly vCAddr1 gCustomFedexCreate
Gui, Add, Edit,  xp yp+24 wp hp ReadOnly vCAddr2 gCustomFedexCreate
Gui, Add, Edit,  xp yp+24 wp hp ReadOnly vCAddr3 gCustomFedexCreate
Gui, Add, Edit,  xp yp+24 wp hp ReadOnly vCZIP gCustomFedexCreate
Gui, Add, Edit,  xp yp+24 wp hp ReadOnly vCCity gCustomFedexCreate
Gui, Add, Edit,  xp yp+24 wp hp ReadOnly vCPhone gCustomFedexCreate
Gui, Add, Edit,  xp yp+24 wp hp ReadOnly vCEmail gCustomFedexCreate
Gui, Add, Edit,  xp yp+24 wp hp ReadOnly vCLink
Gui, Add, Edit,  x136 y208 w202 h21 ReadOnly -TabStop vCCountryName
gui, font
Gui, Add, Checkbox, x56 y448 w90 h30 gEditFields vCheckbox -TabStop, Enable Editing

Gui, Add, Button, x240 y448 w100 h30 gCustomFedex +Disabled -TabStop, Custom FedEx!
Gui, Add, Button, x155 y448 w80 h30 gCCodes +Disabled -TabStop, Country Codes
Gui, Add, Button, x16 y170 w80 h30 gOpenSettings -TabStop, Settings
;Gui, 2:hide
gui, font, s6 bold, Verdana
; Gui, Add, Text, x8 y192 w100 h15 +0x200, Press ESC to close
gui, Tab



; Settings Ini
SettingsLoc = %A_ScriptDir%\Settings.ini

IniRead, filepresentornot,%SettingsLoc%
if filepresentornot =
{

FileAppend,  ; The comma is required in this case.
(
[Settings]
ScreenshotPath=%A_desktop%

[CountryConversion]	add new line - format: School_Code=Fedex_Code
SM=SAMA:IT


[Countries] do not change - these are FEDEX standards
Data=Afghanistan,AF|Albania,AL|Algeria,DZ|American Samoa,AS|Andorra,AD|Angola,AO|Anguilla,AI|Antigua,AG|Argentina,AR|Armenia,AM|Aruba,AW|Australia,AU|Austria,AT|Azerbaijan,AZ|Bahamas,BS|Bahrain,BH|Bangladesh,BD|Barbados,BB|Barbuda (Antigua),BAR:AG|Belarus,BY|Belgium,BE|Belize,BZ|Benin,BJ|Bermuda,BM|Bhutan,BT|Bolivia,BO|Bonaire (Caribbean Netherlands),BON:BQ|Bosnia-Herzegovina,BA|Botswana,BW|Brazil,BR|British Virgin Islands,VG|Brunei,BN|Bulgaria,BG|Burkina Faso,BF|Burundi,BI|Cambodia,KH|Cameroon,CM|Canada,CA|Canary Islands (Spain),CAIS:ES|Cape Verde,CV|Caribbean Netherlands,BQ|Cayman Islands,KY|Chad,TD|Channel Islands (United Kingdom),CHIS:GB|Chile,CL|China,CN|Colombia,CO|Congo,CG|Congo (Democratic Republic of),CD|Cook Islands,CK|Costa Rica,CR|Croatia,HR|Cuba,CU|Curacao,CW|Cyprus,CY|Czech Republic,CZ|Denmark,DK|Djibouti,DJ|Dominica,DM|Dominican Republic,DO|East Timor,TL|Ecuador,EC|Egypt,EG|El Salvador,SV|England (United Kingdom),ENG:GB|Equatorial Guinea,GQ|Eritrea,ER|Estonia,EE|Ethiopia,ET|Faroe Islands,FO|Fiji,FJ|Finland,FI|France,FR|French Guiana,GF|French Polynesia,PF|Gabon,GA|Gambia,GM|Georgia,GE|Germany,DE|Ghana,GH|Gibraltar,GI|Grand Cayman (Cayman Islands),GRCA:KY|Great Britain (United Kingdom),GRBR:GB|Great Thatch Islands (British Virgin Islands),GRTH:VG|Great Tobago Islands (British Virgin Islands),GRTO:VG|Greece,GR|Greenland,GL|Grenada,GD|Guadeloupe,GP|Guam,GU|Guatemala,GT|Guinea,GN|Guyana,GY|Haiti,HT|Honduras,HN|Hong Kong,HK|Hungary,HU|Iceland,IS|India,IN|Indonesia,ID|Iraq,IQ|Ireland,IE|Israel,IL|Italy,IT|Ivory Coast,CI|Jamaica,JM|Japan,JP|Jordan,JO|Jost Van Dyke Islands (British Virgin Islands),JVDI:VG|Kazakhstan,KZ|Kenya,KE|Kiribati,KI|Kuwait,KW|Kyrgyzstan,KG|Laos,LA|Latvia,LV|Lebanon,LB|Lesotho,LS|Liberia,LR|Libya,LY|Liechtenstein,LI|Lithuania,LT|Luxembourg,LU|Macau,MO|Macedonia,MK|Madagascar,MG|Malawi,MW|Malaysia,MY|Maldives,MV|Mali,ML|Malta,MT|Marshall Islands,MH|Martinique,MQ|Mauritania,MR|Mauritius,MU|Mexico,MX|Micronesia,FM|Moldova,MD|Monaco,MC|Mongolia,MN|Montenegro,ME|Montserrat,MS|Morocco,MA|Mozambique,MZ|Namibia,NA|Nauru,NR|Nepal,NP|Netherlands,NL|New Caledonia,NC|New Zealand,NZ|Nicaragua,NI|Niger,NE|Nigeria,NG|Niue,NU|Norman Island (British Virgin Islands),NOIS:VG|Northern Ireland (United Kingdom),NOIR:GB|Northern Mariana Islands,MP|Norway,NO|Oman,OM|Pakistan,PK|Palau,PW|Palestine,PS|Panama,PA|Papua New Guinea,PG|Paraguay,PY|Peru,PE|Philippines,PH|Poland,PL|Portugal,PT|Puerto Rico,PR|Qatar,QA|Reunion,RE|Romania,RO|Rota (Northern Mariana Islands),ROT:MP|Russia,RU|Rwanda,RW|Saba (Caribbean Netherlands),SAB:BQ|Saipan (Northern Mariana Islands),SAI:MP|Samoa,WS|San Marino (Italy),SAMA:IT|Saudi Arabia,SA|Scotland (United Kingdom),SCO:GB|Senegal,SN|Serbia,RS|Seychelles,SC|Singapore,SG|Slovak Republic,SK|Slovenia,SI|Solomon Islands,SB|South Africa,ZA|South Korea,KR|Spain,ES|Sri Lanka,LK|St Barthelemy (Guadeloupe),STBA:GP|St Christopher (Saint Kitts And Nevis),STCH:KN|St Croix Island (U S Virgin Islands),STCR:VI|St Eustatius (Caribbean Netherlands),STEU:BQ|St John (U S Virgin Islands),STJO:VI|St Kitts and Nevis,KN|St Lucia,LC|St Maarten,SX|St Martin,MF|St Thomas (U S Virgin Islands),STTH:VI|St Vincent,VC|Suriname,SR|Swaziland,SZ|Sweden,SE|Switzerland,CH|Syria,SY|Tahiti (French Polynesia),TAH:PF|Taiwan,TW|Tanzania,TZ|Thailand,TH|Tinian (Northern Mariana Islands),TIN:MP|Togo,TG|Tonga,TO|Tortola Island (British Virgin Islands),TOIS:VG|Trinidad and Tobago,TT|Tunisia,TN|Turkey,TR|Turkmenistan,TM|Turks and Caicos Islands,TC|Tuvalu,TV|Uganda,UG|Ukraine,UA|Union Island (St Vincent),UNIS:VC|United Arab Emirates,AE|United Kingdom,GB|United States,US|U.S. Virgin Islands,VI|Uruguay,UY|Uzbekistan,UZ|Vanuatu,VU|Vatican City (Italy),VACI:IT|Venezuela,VE|Vietnam,VN|Wales (United Kingdom),WAL:GB|Wallis and Futuna Islands,WF|Yemen,YE|Zambia,ZM|Zimbabwe,ZW



), %SettingsLoc%

msgbox Settings.ini file created at %SettingsLoc%.	; this fires if 
}
; Else
; msgbox continue normally please.




; IniRead, OutputVar, Filename, Section, Key , Default
IniRead,SSPath,%SettingsLoc%,Settings,ScreenshotPath
if SSPath = ERROR
SSPath = %A_desktop%


IniRead,datastring,%SettingsLoc%,Countries,Data
StringReplace, datastring, datastring, |,`n, All
; maybe add something to prevent errors or something.
;	[Countries]
;	Data=Afghanistan,AF|Albania,AL
; ^^ looks like this ^^



; THIRD Gui
Gui 3:new, -MinimizeBox -MaximizeBox +SysMenu +ToolWindow -Theme
Gui, 3:Add, Text, +0x200, Screenshot location:
Gui, 3:Add, Edit, w400 r1 vSSLocation
Gui, 3:Add, Button, w80 h30 gSelectFolder, Select Folder
Gui, 3:Add, Button, x230 y170 w80 h30 gSaveSettings, Save
Gui, 3:Add, Button, x320 yp w80 h30 gCancelSettings, Close
GuiControl,, SSLocation, %SSPath%




; FOURTH Gui
create4thGUI:
Gui 4:new, -MinimizeBox -MaximizeBox +SysMenu +ToolWindow -Theme
; Gui, 4:Add, Edit, x8 y5 w200 vSearchTerm gSearch
Gui, 4:Add, Text, +0x200, Double click to select.
/*
block commented to avoid using DLL
; ; Gui, 4:Add, ListView, hWnd_hLV1 x8 y28 w345 h346 +checked +Grid +LV0x10000 +LV0x840 -Multi -TabStop gMyListView, Country Name|Code
; ; DllCall("UxTheme.dll\SetWindowTheme", "Ptr", _hLV1, "WStr", "Explorer", "Ptr", 0)
*/
Gui, 4:Add, ListView, w345 h346  +checked +Grid +LV0x10000 +LV0x840 -LV0x4 -Multi -TabStop gMyListView, Country Name|Code
; Gosub, createArray
; Gosub, ccodes
advancedstatus = 0
Progress, off

IniRead,CountryConversionData,%SettingsLoc%,CountryConversion


}
Gosub, info
Return	; first Return




OnError("RaiseE")

RaiseE(exception) {
    msgbox % "Error on line " exception.Line ": " exception.Message "`n"
        , errorlog.txt
    ; return true
}



CountryConvert(code)
{
	global CountryConversionData	; defined in the beginning.
; StringReplace, CCData, CountryConversionData, `n, |, All
CCData := StrReplace(CountryConversionData, "`n", "|")
; msgbox % "CCData: `n" CCData "`n`n ------ `n CountryConversionData: " CountryConversionData

;BUILD CCArray[Row,Column] FROM CountryConversionData
CCArray := {}
MaxRow := 0
MaxCol := 0

Loop, parse, CCData, |	;PARSE AT NEW LINES TO GET RowString
{
	RowIndex := A_Index
	RowString := A_LoopField
	Loop, parse, RowString, =  ;PARSE AT = TO GET Elements
	{
		ColIndex := A_Index
		Element  := A_LoopField
		CCArray[RowIndex,ColIndex] := Element ;BUILD CCArray[row,col] FROM Elements
	}
	If (ColIndex > MaxCol)	;IF ColIndex > MaxCol THEN UPDATE MaxCol
		MaxCol := ColIndex
}



MaxRow := RowIndex
;BUILD args[ColIndex] FROM CCArray[Row2..,Col1..]
MaxRow := MaxRow   		; -1				;REDUCE MaxRow BY 1 TO ACCOUNT FOR COLTITLE ROW
Loop, %MaxRow% 
{
	args := {}
	RowIndex := A_Index   		; +1			; INCREASE RowIndex BY 1 TO ACCOUNT FOR COLTITLE ROW
	Loop, %MaxCol%
	{
		ColIndex := A_Index
		args[ColIndex] := CCArray[RowIndex,ColIndex]
			; msgbox % "found it (ee=ff): " CCArray[2, 1]
		if CCArray[RowIndex,1] = code	; "EE" for example
		{
			; msgbox % "found it (ee=ff): " CCArray[RowIndex, 2]
			Return CCArray[RowIndex,2]	; Return the fedex_code
		}
	
	}
}
}




CustomFedexCreate:
; MsgBox % "ccountry: " CCountry 
Gui, 1:Submit, NoHide



outp := CountryConvert(CCountry)
if outp !=
	GuiControl, Text, Edit2, % outp



Gui, 1:Submit, NoHide
; MsgBox % "ccountry: " CCountry 
mesaj := findcountry(CCountry)
ControlSetText, Edit12, % mesaj, ahk_id %_2hwnd%
; MsgBox % "ccountry: " CCountry "`nE_CountryField: " E_CountryField
Return

CustomFedex:
Gui, 1:Submit, NoHide

mesaj := findcountry(CCountry)
ControlSetText, Edit12, % mesaj, ahk_id %_2hwnd%
if mesaj = 
{
; MsgBox edit12 empty %CCountry%
CCountry := ""
ControlSetText, Edit2, , ahk_id %_2hwnd%
}

E_CountryField := LC_UriEncode(CCountry)
E_NameField := LC_UriEncode(CName)
E_Addr1Field := LC_UriEncode(CAddr1)
E_Addr2Field := LC_UriEncode(CAddr2)
E_CityField := LC_UriEncode(CCity)
E_ZIPField := LC_UriEncode(CZIP)
E_PhoneField := LC_UriEncode(CPhone)
E_emailField := LC_UriEncode(CEmail)

customfedexStart := "https://www.fedex.com/shipping/shipEntryAction.do?origincountry=us&locallang=us&urlparams=us&toData.addressData.countryCode="

customfedexlink := customfedexStart . E_CountryField . "&toData.addressData.contactName=" 
. E_NameField . "&toData.addressData.addressLine1=" . E_Addr1Field . "&toData.addressData.addressLine2=" . E_Addr2Field . "&toData.addressData.zipPostalCode=" . E_ZIPField . "&toData.addressData.city=" . E_CityField . "&toData.addressData.phoneNumber=" . E_PhoneField . "&notificationData.recipientNotifications.email=" . E_emailField .  "&notificationData.recipientNotifications.tenderedNotificationFlag=true&notificationData.recipientNotifications.exceptionNotificationFlag=true&notificationData.recipientNotifications.deliveryNotificationFlag=true&psdData.weightUnitOfMeasure=LBS&psdData.mpsRowDataList[0].weight=0.5&commodityData.documentShipping=true&commodityData.shipmentPurposeCode=7&commodityData.totalCustomsValue=1&commodityData.documentDescriptionCode=25"

ControlSetText, Edit11, % customfedexlink, ahk_id %_2hwnd%

; MsgBox %customfedexlink%
Run, chrome.exe "%customfedexlink%"
Return



theloading:
Progress, b w200, , Loading,
loop, 100
{
Progress, %a_index% ; Set the position of the bar to 50%.
}
Return


;;;;;;;;;;;	BEGIN Settings GUI Events ;;;;;;;;;;;
{
OpenSettings:
Gui, 3:Show, w420 h220, Settings	; x1000 y385
Return

SelectFolder:
FileSelectFolder, OutputVar, , 3
if OutputVar =
{
    sleep, 50
	; MsgBox, You didn't select a folder.
}
else
{
	OutputVar := OutputVar "\"
	GuiControl,, SSLocation, %OutputVar%
    ; MsgBox, You selected folder "%OutputVar%".
}
return

SaveSettings:
GuiControlGet, SSLocation
stringLeft, var_name, SSLocation, 1
if var_name = C
{
	IniWrite,%SSLocation%,%SettingsLoc%,Settings,ScreenshotPath
	sleep, 200
	Gui, 3:Hide
}
else
{
	SetTimer, ChangeDialogButton, 5
	msgbox, 52, Caution, Saving in network drives is not recommended.`n`nContinue saving in this location?
	IfMsgBox No
	{
		GuiControl,, SSLocation, %SSPath%
		ControlFocus, Edit1
		
		Return
	}
	IfMsgBox Yes
	{
		IniWrite,%SSLocation%,%SettingsLoc%,Settings,ScreenshotPath
		sleep, 20
		Gui, 3:Hide
	}
}
return

ChangeDialogButton: 
IfWinNotExist, Caution
    return  ; Keep waiting.
SetTimer, ChangeDialogButton, Off 
; WinActivate 
ControlSetText, Button1, &Save here
ControlSetText, Button2, &Change 
return


CancelSettings:
Gui, 3:Hide
return
}
;;;;;;;;;;;	END Settings GUI Events ;;;;;;;;;;;



Info:	; expands the GUI with student info

if advancedstatus = 0
{
	Gui, 1:Show, w360 h486
	GuiControl, Choose, SysTabControl321, Main
	GuiControl,, Button4, Less ⮃
	GuiControl +BackgroundFFFFFF, Button4
	GuiControl +BackgroundFF9977, SysTabControl321
	advancedstatus = 1
	; gosub, recolor
}
Else
{
	Gui, 1:Show, w360 h164
	GuiControl, Choose, SysTabControl321, Out
	GuiControl,, Button4, More ⮃
	advancedstatus = 0
}

Update:	; update fields in the gui2 as the DDL selection changes, see: gUpdate in the gui line.
SplitDDL(a,EduID) ; split function get EduID from DDL selection
Student := StuInfo(EduID)	; use EduID from DDL selection to bring student info.
; MsgBox % "fxfedexlink: " Student.Link

student_country := Student.Country



outp := CountryConvert(student_country)
if outp !=
{	; GuiControl, Text, Edit2, % outp
	student_country := outp
	; msgbox update`noutp: %outp% `nstudent_country: %student_country%
}



ControlSetText,,% Student.ID, ahk_id %_EduID%
ControlSetText, Edit2, % student_country, ahk_id %_2hwnd%
mesaj := findcountry(CCountry)
; MsgBox % "mesaj: " mesaj
ControlSetText, Edit12, % mesaj, ahk_id %_2hwnd%
ControlSetText, Edit3, % Student.Name, ahk_id %_2hwnd%
ControlSetText, Edit4, % Student.Addr1, ahk_id %_2hwnd%
ControlSetText, Edit5, % Student.Addr2, ahk_id %_2hwnd%
ControlSetText, Edit6, % Student.Addr3, ahk_id %_2hwnd%
ControlSetText, Edit7, % Student.ZIP, ahk_id %_2hwnd%
ControlSetText, Edit8, % Student.City, ahk_id %_2hwnd%
ControlSetText, Edit9, % Student.Phone, ahk_id %_2hwnd%
ControlSetText, Edit10, % Student.email, ahk_id %_2hwnd%
ControlSetText, Edit11, % Student.link, ahk_id %_2hwnd%

addr1 := Student.Addr1
addr2 := Student.Addr2
addr3 := Student.Addr3
addr1len := StrLen(Addr1)
addr2len := StrLen(Addr2)
addr3len := StrLen(Addr3)
firstaddrlen := addr1len + addr2len
firstaddrline = %Addr1%, %Addr2%
secondaddrlen := addr2len + addr3len
secondaddrline = %Addr2%, %Addr3%

ControlSetText, Static4, Address 1 (%addr1len%), ahk_id %_2hwnd%
ControlSetText, Static5, Address 2 (%addr2len%), ahk_id %_2hwnd%
ControlSetText, Static6, Address 3 (%addr3len%), ahk_id %_2hwnd%

; if (addr1len > 35 or addr2len > 35 or addr3len > 35)
; {
	; ; msgbox Address lines too long (no more than 35 chars pls)`nFirst line (%addr1len%): %Addr1%`nSecond line (%addr2len%): %Addr2%`nThird line (%addr3len%): %Addr3%
; }
if (addr1len < 36 and addr2len < 36 and addr3len = 0)
{
	; msgbox  ideal
	; ideal condition, do nothing.
	; send Addr1 and Addr2
}
Loop 3
{
	textnumber := a_index + 3
	if addr%a_index%len > 35	; msgbox %a_index%. line is too long!
	{
		guicontrol, +cRed +Redraw, Static%textnumber%
		; guicontrol, +cRed +Redraw, Edit%textnumber%
		; edit lines are broken in the next loop
	}
	else
	{
		guicontrol, +cBlack +Redraw, Static%textnumber%
		; guicontrol, +cBlack +Redraw, Edit%textnumber%
	}
}
loop 3
{
	textnumber := a_index + 3
	if addr3len != 0
	{
		if (firstaddrlen < 34 and addr3len < 34)
		{
			; msgbox % firstaddrlen " `n " addr3len
			guicontrol, +cGreen +Redraw, Edit%textnumber%
		}
		else if (firstaddrlen > 33 and secondaddrlen < 34)
		{
			; msgbox % firstaddrlen " `n " secondaddrlen
			guicontrol, +cGreen +Redraw, Edit%textnumber%
		}
		else if (firstaddrlen > 33 and secondaddrlen > 33)
		{
			; msgbox % firstaddrlen " `n " secondaddrlen
			guicontrol, +cRed +Redraw, Edit%textnumber%
		}
		else
		{
			guicontrol, +cBlack +Redraw, Edit%textnumber%
		}
	}
	else
		guicontrol, +cBlack +Redraw, Edit%textnumber%
}

Return



GuiDropFiles:
extension := SubStr(A_GuiEvent, -3)
if SubStr(A_GuiEvent, -3) = "xlsx"	; change according to actual report
{
	; MsgBox, Opening report:`n`r%A_GuiEvent%
	gosub, theloading
	GuiControl,, List, |
	FilePath := A_GuiEvent
	oWorkbook := ComObjGet(FilePath) ; access Workbook object
	IDcolRange := oWorkbook.Sheets(1).Range("A1:A300")
	gosub, load2
	Progress, off
	MsgBox, Report loaded:`n`r%A_GuiEvent%`nReady to go.
}
else
; MsgBox,278544, error, drag and drop xlsx file. ("%extension%" is invalid.)
MsgBox,262160, Error, Drag and drop an Excel (.xlsx) file. ("%extension%" is invalid.)
; MsgBox,4112, error, drag and drop xlsx file. ("%extension%" is invalid.)
Return




load:
FileSelectFile, SelectedFile, 3, , Open a file, Excel Documents (*.xlsx; *.xls)	;	rename to whatever the report's name is.
if SelectedFile =
{
	Return
    ; MsgBox, The user didn't select anything.
}
else
{
	MsgBox, The user selected the following:`n%SelectedFile%
	GuiControl,, List, |
	FilePath := SelectedFile ; example path
	oWorkbook := ComObjGet(FilePath) ; access Workbook object
	IDcolRange := oWorkbook.Sheets(1).Range("A1:A300")
	gosub, load2
	MsgBox, Report loaded:`n`r%A_GuiEvent%`nReady to go.
}
Return

Load2:	; next gui Button

firstrow = 3		; the row the list starts
EduIDCol = A		; the column the EduID resides
lastrow:= oWorkbook.Sheets(1).Range["A" . firstrow].End(-4121).Address[0,0]
lastrowno := RegExReplace(lastrow, "\D")
loopcounter := lastrowno - firstrow + 1
; MsgBox 	number of students:  %lastrowno%`r`nCreating a list of %lastrowno% fedex links
list := ""
loop, %loopcounter%	; loop number of students in the list
{
	EduID := oWorkbook.Sheets(1).Range(EduIDCol . firstrow).Value
	Student := StuInfo(EduID)
	name := Student.Name

	list .= "|" name " — " EduID
	firstrow+=1	; off to the next row!!
}
list := StrReplace(list, "||", "|")
GuiControl,, List, ❤Student names❤||%list%
; msgbox %list%
PostMessage, 0x014E, 0, 0,,ahk_id %hwndvar% ; select the first one

Return


findcountry(code)
{
	global DataString	; defined in the beginning.

;BUILD ARRAY[Row,Column] FROM DataString
Array     := {}
MaxRow := 0
MaxCol  := 0

Loop, parse, DataString, `n  					;PARSE DataString AT NEW LINES TO GET RowString
{
	RowIndex := A_Index
	; msgbox % RowIndex
	RowString := A_LoopField
	; msgbox % RowString
	Loop, parse, RowString, CSV 				;PARSE RowString AT COMMAS TO GET Elements
	{
		ColIndex := A_Index
		; msgbox % ColIndex
		Element  := A_LoopField
		; msgbox % Element
		Array[RowIndex,ColIndex] := Element 	;BUILD Array[row,col] FROM Elements
		; msgbox % Array[2].1
	}
	If (ColIndex > MaxCol)					;IF ColIndex > MaxCol THEN UPDATE MaxCol
		MaxCol := ColIndex
}
	
MaxRow := RowIndex
;BUILD args[ColIndex] FROM Array[Row2..,Col1..]
MaxRow := MaxRow   		; -1				;REDUCE MaxRow BY 1 TO ACCOUNT FOR COLTITLE ROW
Loop, %MaxRow% 
{
	args := {}
	RowIndex := A_Index   		; +1			; INCREASE RowIndex BY 1 TO ACCOUNT FOR COLTITLE ROW
	Loop, %MaxCol%
	{
		ColIndex := A_Index
		args[ColIndex] := Array[RowIndex,ColIndex]
		if Array[RowIndex,2] = code
		{
			Return Array[RowIndex,1]	; Return the country name 
		}
	
	}
}
args =
; Array =
}


MyListView:
if A_GuiEvent = DoubleClick
{
    LV_GetText(LV_countrycode, A_EventInfo, 2)  ; Get the text from the row's first field.
    LV_GetText(LV_countryname, A_EventInfo, 1)  ; Get the text from the row's first field.
	ControlSetText, Edit2, % LV_countrycode, ahk_id %_2hwnd%
	ControlSetText, Edit12, % LV_countryname, ahk_id %_2hwnd%
	Gui, 1:Show
	Gui, 4:hide
}
return

ccodes:		; destroy and recreate the fourth gui w/ listbox of countries
; Gui, 4:Destroy	; I need to delete LV instead of destroying this.
; gosub, create4thGUI	; create fourth gui
advancedstatus = 1	; fix the more/less thingy caused by the previous line 
gosub, createArray
Gui, 4:Show, w360 h381, Country Codes	; x1000 y385
Return


createArray:
Gui, 4:Default
LV_Delete()	; delete all rows first.
;BUILD ARRAY[Row,Column] FROM DataString
Array     := {}
MaxRow := 0
MaxCol  := 0
Loop, parse, DataString, `n  					;PARSE DataString AT NEW LINES TO GET RowString
{
	RowIndex := A_Index
	RowString := A_LoopField
	Loop, parse, RowString, CSV 				;PARSE RowString AT COMMAS TO GET Elements
	{
		ColIndex := A_Index
		Element  := A_LoopField
		Array[RowIndex,ColIndex] := Element 	;BUILD Array[row,col] FROM Elements
	}
	If (ColIndex > MaxCol)					;IF ColIndex > MaxCol THEN UPDATE MaxCol
		MaxCol := ColIndex
}
MaxRow := RowIndex
;BUILD args[ColIndex] FROM Array[Row2..,Col1..]
MaxRow := MaxRow   		; -1				;REDUCE MaxRow BY 1 TO ACCOUNT FOR COLTITLE ROW
Loop, %MaxRow% 
{
	args := {}
	RowIndex := A_Index   		; +1			; INCREASE RowIndex BY 1 TO ACCOUNT FOR COLTITLE ROW
	Loop, %MaxCol%
	{
		ColIndex := A_Index
		args[ColIndex] := Array[RowIndex,ColIndex]
	}
	;ADD ROW TO ListView USING args*
	LV_Add("",args*)
	; MsgBox % args[1]
}
args =
; Array =
; MsgBox % array[52].2
; MsgBox % array[52,1]
LV_ModifyCol()
Return


StuInfo(ID)
{
global
; FilePath := A_Desktop "\fedexList.xlsx" ; example path
; oWorkbook := ComObjGet(FilePath) ; access Workbook object
; IDcolRange := oWorkbook.Sheets(1).Range("A1:A300")
; msgbox % isobject(IDcolRange)
; MsgBox % "file:" FilePath
f := IDcolRange.Find(ID,,, xlWhole:=true)
; MsgBox % "f:" f.value
EduIDField := f.value		; the anchor ; rest is set by offset
NameFieldLast := f.offset(0,1).value ; rest is set by offset
NameFieldFirst := f.offset(0,2).value 
NameField := NameFieldFirst " " NameFieldLast
CountryField := f.offset(0,3).value	
Addr1Field := f.offset(0,4).value
Addr2Field := f.offset(0,5).value
Addr3Field := f.offset(0,6).value
CityField := f.offset(0,7).value
ZIPFieldRaw := f.offset(0,8).value
ZIPFieldArray := StrSplit(ZIPFieldRaw, ".000000")	; excel adds, I remove
ZIPField := ZIPFieldArray[1]
PhoneFieldRaw := f.offset(0,9).value
PhoneFieldArray := StrSplit(PhoneFieldRaw, ".000000")
PhoneField := PhoneFieldArray[1]
emailField := f.offset(0,10).value

; MsgBox % Addr1Field

outp := CountryConvert(CountryField)
if outp !=
{	; GuiControl, Text, Edit2, % outp
	CountryField := outp
	; msgbox % "outp: " outp "`ncountryField: " CountryField
}

E_CountryField := LC_UriEncode(CountryField)
E_NameField := LC_UriEncode(NameField)
E_Addr1Field := LC_UriEncode(Addr1Field)
E_Addr2Field := LC_UriEncode(Addr2Field)
E_Addr3Field := LC_UriEncode(Addr3Field)
E_CityField := LC_UriEncode(CityField)
E_ZIPField := LC_UriEncode(ZIPField)
E_PhoneField := LC_UriEncode(PhoneField)
E_emailField := LC_UriEncode(emailField)

; if country code is non existent in fedex codes, set it empty and don't send it to fedex.com 
mesaj := findcountry(CountryField)	
; MsgBox % mesaj "`n" E_CountryField
if mesaj = 
{
E_CountryField := ""
CountryField := ""
}
; MsgBox % mesaj "`n" E_CountryField "`n" CountryField 


fedexStart := "https://www.fedex.com/shipping/shipEntryAction.do?origincountry=us&locallang=us&urlparams=us&toData.addressData.countryCode="

fedexlink := fedexStart . E_CountryField . "&toData.addressData.contactName=" 
. E_NameField . "&toData.addressData.addressLine1=" . E_Addr1Field . "&toData.addressData.addressLine2=" . E_Addr2Field . "&toData.addressData.zipPostalCode=" . E_ZIPField . "&toData.addressData.city=" . E_CityField . "&toData.addressData.phoneNumber=" . E_PhoneField . "&notificationData.recipientNotifications.email=" . E_emailField .  "&notificationData.recipientNotifications.tenderedNotificationFlag=true&notificationData.recipientNotifications.exceptionNotificationFlag=true&notificationData.recipientNotifications.deliveryNotificationFlag=true&psdData.weightUnitOfMeasure=LBS&psdData.mpsRowDataList[0].weight=0.5&commodityData.documentShipping=true&commodityData.shipmentPurposeCode=7&commodityData.totalCustomsValue=1&commodityData.documentDescriptionCode=25"


Test := {Link: fedexlink, ID: EduIDField, Name: NameField, Country: CountryField, Addr1: Addr1Field, Addr2: Addr2Field, Addr3: Addr3Field, City: CityField, ZIP: ZIPField, Phone: PhoneField, email: emailField, e_name: E_NameField, e_Country: e_CountryField, e_Addr1: e_Addr1Field, e_Addr2: e_Addr2Field, e_Addr3: e_Addr3Field, e_City: e_CityField, e_ZIP: e_ZIPField, e_Phone: e_PhoneField, e_email: e_emailField}

return Test
}

EditFields:
if check
{
	GuiControl, Disable, button8
	GuiControl, Disable, button9
	time := 2
	loop 9
	{
		GuiControl, +ReadOnly -readwrite, Edit%time%
		; GuiControl, +ReadOnly -readwrite, Edit1	; repeat for all
		time += 1
	}
}
Else
{
	GuiControl, Enable, button8
	GuiControl, Enable, button9
	time := 2
	loop 9
	{
		GuiControl, +readwrite -ReadOnly, Edit%time%
		; GuiControl, +readwrite -ReadOnly, Edit1
		time += 1
	}
}
check := !check
Return

fedex:
gosub, closechrome
SplitDDL(stringa1,stringa2)
fedexlink := StuInfo(stringa2).link
; msgbox %stringa1% `r`n %stringa2%
; Run, chrome.exe "%fedexlink%"


; Run, chrome.exe --profile-directory="new" --app="%fedexlink%"
; BlockInput On	; doesn't work without admin rights. oh well.
; ; sleep, 5000javascript:document.getElementBode").value=25;

; winwait, FedEx Ship Manager - Create a Shipment - Google Chrome
; ; winwait, FedEx Ship Manager - Create a Shipment
; sleep, 1000
; SendInput ^l
; sleep, 50
; SendInput javascript:document.getElementById("commodityData.documentDescriptionCode").value=25;{enter}
; BlockInput Off
Gosub, newchrome
Return


; #space::Run, chrome.exe --profile-directory="Default" --app="http://example.com/"

SplitDDL(ByRef stringa1, ByRef stringa2) ; splits the selected item into variables: name -- EduID
{
	; example: SplitDDL(a,b)
	GuiControlGet, List
	string_array := StrSplit(List, " — ")
	stringa1 := string_array[1] ; name
	stringa2 := string_array[2] ; EduID
}

next:	; next gui Button
ControlSend, ComboBox1, {down}
; PostMessage, 0x014E, %nextitem%, 0,,ahk_id %hwndvar%
Return

prev:	; previous gui Button
ControlSend, ComboBox1, {up}
Return



LC_UrlEncode(Url) { ; keep ":/;?@,&=+$#."
	return LC_UriEncode(Url, "[0-9a-zA-Z:/;?@,&=+$#.]")
}
LC_UriEncode(Uri, RE="[0-9A-Za-z]") {
	Res:=""
	VarSetCapacity(Var, StrPut(Uri, "UTF-8"), 0), StrPut(Uri, &Var, "UTF-8")
	While Code := NumGet(Var, A_Index - 1, "UChar")
		Res .= (Chr:=Chr(Code)) ~= RE ? Chr : Format("%{:02X}", Code)
	Return, Res
}

down::
IfWinActive, ahk_id %_2hwnd%
{
	Send {tab}
}
Else
	Send	{down}
Return
up::
IfWinActive, ahk_id %_2hwnd%
{
	Send +{tab}
}
Else
	Send	{up}
Return



3GuiEscape:
Gui 3:hide
return


GuiClose:
ExitApp

^r::
; Gosub, closechrome
Reload
Sleep, 500
msgbox,,, DIDN'T WORK
return




; enter chrome.ahk path and FEDEX login info.
#Include .\Chrome ahk_v1.2\Chrome.ahk
FEDEXusername = ""
FEDEXpassword = ""



newchrome:

; --- Create a new Chrome instance ---

FileCreateDir, ChromeProfile
ChromeInst := new Chrome("ChromeProfile")

if (Chromes := Chrome.FindInstances())
	ChromeInst := {"base": Chrome, "DebugPort": Chromes.MinIndex()}
else
	ChromeInst := new Chrome(ProfilePath)




; --- Connect to the page ---

if !(PageInst := ChromeInst.GetPage())
{
	MsgBox, Could not retrieve page!
	ChromeInst.Kill()
}
else
{
	PageInst.Call("Page.navigate", {"url": "https://www.fedex.com/sed"})
	PageInst.WaitForLoad()
	; --- Perform JavaScript injection ---
	username = document.getElementsByName("username")[0].value = FEDEXusername
	password = document.getElementsByName("password")[0].value = FEDEXpassword
	login = document.getElementsByName("login")[0].click();

	PageInst.Evaluate(password)
	PageInst.Evaluate(username)
	PageInst.Evaluate(login)
	

	
	
	
	
	
	PageInst.WaitForLoad()
	PageInst.Call("Page.navigate", {"url": fedexlink})
	; ; testing: 
	; PageInst.Call("Page.navigate", {"url": ".\Desktop\FedEx Ship Manager - Print Your Label(s).html"})
	
	; MsgBox country selected: %E_CountryField%
	if E_CountryField = US
	{
	; MsgBox US selected: %E_CountryField%
	JS1 := "document.getElementById('psdData.serviceType').value='Standard Overnight'"
	JS2 := ""
	JS3 := ""
	timetoloop := 1
	}
	Else
	{
	; MsgBox Non US country: %E_CountryField%
	JS1 := "document.getElementById('commodityData.documentDescriptionCode').value=25;"
	JS2 := "document.getElementById('commodityData.totalCustomsValue').value=1;"
	JS3 := "document.getElementById('commodityData.totalCustomsValueCurrencyCode').value='USD';"
	timetoloop := 3
	}
	
	PageInst.WaitForLoad()
	
	Loop %timetoloop%
	{		
		if (JS%a_index% == "" || ErrorLevel)
			break
		
		try
			Result := PageInst.Evaluate(JS%a_index%)
		catch e
		{
			MsgBox, % "Exception encountered in " e.What ":`n`n"
			. e.Message "`n`n"
			. "Specifically:`n`n"
			. Chrome.Jxon_Dump(Chrome.Jxon_Load(e.Extra), "`t")
			
			continue
		}
		
		; MsgBox, % "Result:`n" Chrome.Jxon_Dump(Result, "`t")
	}

	
	
	loop	; this loop waits for the ship button to be pressed.
	{
		currentpageurl := PageInst.Evaluate("window.location.href").value
		
		if currentpageurl = https://www.fedex.com/shipping/shipAction.handle?method=doContinue
		{
			; MsgBox % "Label created, current url:`n" currentpageurl
			gosub, printingjob
			Break
			Return
		}
		Else
		{
			; MsgBox wrong url. exiting thread.
			; this needs to stay emtpy.
		}
		sleep, 1000
	}
	

}

Return


closechrome:
; --- Close the Chrome instance ---

; try
	PageInst.Call("Browser.close") ; Fails when running headless
; catch
	ChromeInst.Kill()
PageInst.Disconnect()
Return







; screenshot the label print page.
; ^!p::
printingjob:
	; resize the window.
	WinActivate, FedEx Ship Manager - Print Your Label(s) - Google Chrome
	WinWait, % "FedEx Ship Manager - Print Your Label(s) - Google Chrome"
	WinTitle := "FedEx Ship Manager - Print Your Label(s) - Google Chrome"
	SysGet, MonitorWorkArea, MonitorWorkArea
	PostMessage, 0x112, 0xF120,,, % WinTitle,   ; 0x112 = WM_SYSCOMMAND, 0xF120 = SC_RESTORE
	WinMove, % WinTitle,,-7,0, MonitorWorkAreaRight/2+14, MonitorWorkAreaBottom+7


	; MsgBox is it on the left?
	; PageInst.WaitForLoad()
	
	; url of current tab.
	currentpageurl := PageInst.Evaluate("window.location.href").value
	
	if currentpageurl = https://www.fedex.com/shipping/shipAction.handle?method=doContinue
	{
		sleep, 50
		; MsgBox % currentpageurl
	}
	Else
	{
		SetTimer, ChangeButtonNames, 50
		MsgBox, 4117, An error occured, If you think this is a mistake, retry.
		ifmsgbox Retry
			goto, printingjob
		else
		{
			msgbox,0,Continue the process manually, 1. Copy the tracking number.`n2. Take a screenshot and save it.`n3. Print the label.
			Return
		}
	}
	
	trackingnumber := PageInst.Evaluate("document.getElementById('label.trackingNumber').innerText").value
	
	PostMessage, 0x112, 0xF120,,, % WinTitle,   ; 0x112 = WM_SYSCOMMAND, 0xF120 = SC_RESTORE
	WinMove, % WinTitle,,-7,0, MonitorWorkAreaRight/2+14, MonitorWorkAreaBottom+7
	; click the print button.
	PageInst.Evaluate("document.getElementById('button.print').click()")
	
	trackingnumber := RegExReplace(trackingnumber, "\D")
	length := strlen(trackingnumber)
	if length = 12
	{
		Clipboard := trackingnumber
		clipwait, 1
		soundbeep, 500, 20
		MsgBox FedEx number copied, Length: %length%`n`r%trackingnumber%
	}
	else
	{
		trackingnumber :=
	}
	

	; msgbox % "this could be broken, watch closely"
	


	Base64PDF := PageInst.Call("Page.captureScreenshot").data
	; Convert to a normal binary PDF
	Size := Base64_Decode(BinaryPDF, Base64PDF)
	; Write the binary PDF to a file
	screenshotname = %NameFieldLast%_%NameFieldFirst% %EduIDField% FEDEX%trackingnumber%.png
	 
	StringRight, var_name, SSPath, 1
	if var_name = "\"
		SSPath := SubStr(SSPath,1,StrLen(SSPath)-1)

	FileOpen(SSPath "\" screenshotname, "w").RawWrite(BinaryPDF, Size)
	MsgBox,,, % "Screenshot saved as: " screenshotname "`n", 5
	; Open the file
	; Run, %screenshotname%, , Min
	
	
	
	; PageInst.Evaluate("document.getElementById('button.newShipment').click()")
; Return

; ^+o::
	MsgBox Printing...
	sleep, 1000
	
	WinGetTitle, Title, A
	PixelGetColor, secondpageNoneGray, 378, 766	;0x595652 gray
	PixelGetColor, pagesAllselectedGray, 150, 307 ;0x666666  gray selected
	PixelGetColor, PrintButtonblue, 219, 173	;0xFFFFFF
	
	; MsgBox % secondpageNoneGray "0x595652`n" pagesAllselectedGray "0x666666`n" PrintButtonblue "0xFFFFFF"
	
	If (PrintButtonblue = 0xFFFFFF and secondpageNoneGray = 0xFFFFFF and pagesAllselectedGray = 0x666666)
	{
		PixelGetColor, PrintButtonblue, 219, 173	;0xFFFFFF
		if (Title = "FedEx Ship Manager - Print Your Label(s) - Google Chrome" and PrintButtonblue = 0xFFFFFF)
		{
			; msgbox, %PrintButtonblue%
			;Return
			MouseClick, left, 235, 348
			MouseClick, left, 235, 348, 3
			sleep, 50
			SendInput 1{tab}2
			; SendInput 1{tab}2
			sleep, 300
			; MsgBox SendInput {enter}
			MsgBox,,, You may click print.,3
		}
		PrintButtonblue =
	}
	
Return

ChangeButtonNames: 
IfWinNotExist, An error occured
    return  ; Keep waiting.
SetTimer, ChangeButtonNames, Off 
WinActivate 
ControlSetText, Button2, &Continue 
return

Base64_Decode(ByRef Out, ByRef In) {
	DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &In, "UInt", StrLen(In)
	, "UInt", 0x1, "Ptr", 0, "UInt*", OutLen, "Ptr", 0, "Ptr", 0)
	VarSetCapacity(Out, OutLen)
	DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &In, "UInt", StrLen(In)
	, "UInt", 0x1, "Str", Out, "UInt*", OutLen, "Ptr", 0, "Ptr", 0)
	return OutLen
}
