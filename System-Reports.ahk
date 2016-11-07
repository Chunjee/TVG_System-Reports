;/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\
; Description
;\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/
; Launches the system report builder with exscalated privileges and then displays the excel output to the user

#NoEnv
#NoTrayIcon
#SingleInstance Force

#Include %A_ScriptDir%\Functions
#Include util_misc.ahk

The_VersionName = v1.0.0

;Read and setup options
FileRead, MemoryFile, %A_ScriptDir%\config.txt
UserName := Fn_QuickRegEx(MemoryFile, "username:(.+)")
Password := Fn_QuickRegEx(MemoryFile, "password:(.+)")
RootDir := Fn_QuickRegEx(MemoryFile, "rootdir:(.+)")
EmailRecipient := Fn_QuickRegEx(MemoryFile, "email:(.+)")

FormatTime, CurrentDate,, MMM/dd/yy
EmailSubject = System Reports %CurrentDate%

;Set the credentials
RunAs, %UserName%, %Password%, TVGOPS

;Run the actual report
if (A_IsCompiled) {
	RunWait, % RootDir . "\Background-System-Report.exe",, UseErrorLevel
} else {
	RunWait, %A_ScriptDir%\System-Report.exe,, UseErrorLevel
}

;find the most recent report excel
The_ExitCode := ErrorLevel
ExcelPath := Fn_FindNewestExcel(The_ExitCode)

;Save Excel
oExcel := ComObjCreate("Excel.Application")
oExcel.Workbooks.Open(ExcelPath) ;open the existing file
oExcel.Visible := True
oExcel.ActiveWorkbook.saved := True


;show the e-mail gui only after the report has been made
GUI()
Return

SendReport:
oExcel.ActiveWorkbook.SendMail(EmailRecipient, EmailSubject)
Return

GuiClose:
ExitApp






;/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\
; Functions
;\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/


Fn_FindNewestExcel(para_ExitCode)
{ ;probably more dependable if we open the current days folder and newest file
global ;needs access to "RootDir"
	if (para_ExitCode = "") {
		Msgbox, There was a problem receiving the exitcode from the backend. `n`nPlease run again and report this problem.
		ExitApp
	}
	;build current time and use to explore daily folder
	FormatTime, l_CurrentYear,, yyyy
	FormatTime, l_CurrentMonth,, MMMM
	FormatTime, l_CurrentDay,, dd
	l_FileDir= %RootDir%\Reports\Archive\%l_CurrentYear%\%l_CurrentMonth%\%l_CurrentDay%\
	Loop, Files, %l_FileDir%*
	{
		last := A_LoopFileFullPath
	}
	return last
}


GUI()
{
global
Gui +AlwaysOnTop

Gui, Font, s14 w70, Arial
Gui, Add, Text, x2 y0 w280 h40 +Center, System Reports
Gui, Font, s10 w70, Arial

Gui, Show, x127 y87 h80 w288, System Reports

Gui, Add, Button, x2 y40 w284 h30 gSendReport, E-mail Report

Return
}