;/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\
; Description
;\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/
; Reads Registry of all systems in \Data\AllSystems.txt and writes to \\NAS\WOG_System_Reports\Data\Active\
; 


;~~~~~~~~~~~~~~~~~~~~~
;Compile Options
;~~~~~~~~~~~~~~~~~~~~~
SetBatchLines -1 ;Go as fast as CPU will allow
StartUp()
The_VersionName = v2.16.0

;Dependencies
#Include %A_ScriptDir%\Functions
#Include util_misc.ahk
#Include util_arrays.ahk

;/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\
;PREP AND STARTUP
;\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/

FormatTime, G_Day, , LongDate
FormatTime, G_Time, , Time

;create GUI and progress bar
GUI()
GUI_DiableAllButtons()
Fn_GUI_UpdateProgress(1)


;This one is an absolute location
A_SharedDir = \\tvgops\pdxshares\wagerops\Tools\System-Reports
;Get all the temp directories prepared/cleared
All_Systems = %A_ScriptDir%\AllSystems.txt

Dir_ActiveConnections = %A_ScriptDir%\Data\Active
Dir_MasterConnections = %A_ScriptDir%\Data\Master
Dir_ArchiveConnections = %A_ScriptDir%\Data\Active\Archived


Path_Temp = %A_ScriptDir%\Data\temp
Path_Temp_Discrepant = %A_ScriptDir%\Data\temp\discrepant_servers
Path_Temp_TGPs = %A_ScriptDir%\Data\temp\TGP_Connections


;Delete Temp
FileRemoveDir, %Path_Temp% , 1

;Delete and Active Connections Not mentioned in AllSystems.txt
Sb_ClearOutofProductionMachines()

Fn_CreateTempDir(Dir_ActiveConnections)
Fn_CreateTempDir(Dir_ArchiveConnections)
Fn_CreateTempDir(Path_Temp)
Fn_CreateTempDir(Path_Temp_TGPs)
Fn_CreateTempDir(Path_Temp_Discrepant)




;/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\
; Main
;\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/

Software_Array := []
TxtDocument_Array := []
The_Ping_Results_Array := [[x],[y]]
Fn_Txt2Array(TxtDocument_Array, A_ScriptDir . "\AllSoftware.txt")

Loop, Read, %All_Systems%
{
	TotalFiles := A_Index
}

X = 0
;Read AllSystems text file
Loop, Read, % All_Systems
{
Fn_GUI_UpdateProgress(A_Index, TotalFiles)

;Remove all spaces from read line
StringReplace, A_LoopFileNameNoSpace, A_LoopReadLine, %A_SPACE%, , All
	;Archive file and Skip any Commented systems
	IfInString, A_LoopFileNameNoSpace, `;
	{
		The_CommentedServerName := Fn_QuickRegEx(A_LoopFileNameNoSpace,"([a-zA-Z]*\d{2})")
		FileMove, %Dir_ActiveConnections%\%The_CommentedServerName%.txt, %Dir_ArchiveConnections%\%The_CommentedServerName%.txt, 1
		Continue
	}
	;Skip if line is blank
	If (A_LoopFileNameNoSpace = "")
	{
		Continue
	}
	
	The_SystemName := Fn_QuickRegEx(A_LoopFileNameNoSpace,"(\D+\d{2})")
	;Send to Gui if valid system name
	If (The_SystemName != "") {
		guicontrol, Text, GUI_CurrentSystem, % The_SystemName
	} else {
		Continue
	}
	
	;Insert the machine into an array for later
	X += 1
	Software_Array[X, "SystemName"] := The_SystemName
	
	;Ping the system and skip if there is no reply.
	The_PingResult := Fn_Ping(The_SystemName)
	;runwait, %comspec% /c ping %The_SystemName% -n 1 -w %The_PingTime%,,UseErrorLevel Hide
		;What to do if ping is success or failure
		If (The_PingResult = 1)
		{
			;success
			FileDelete, %Dir_ActiveConnections%\%The_SystemName%.txt
			The_Ping_Results_Array[A_Index,1] := The_SystemName
			The_Ping_Results_Array[A_Index,2] := The_PingResult
		} else {
			;failure
			The_Ping_Results_Array[A_Index,1] := The_SystemName
			The_Ping_Results_Array[A_Index,2] := The_PingResult
			Continue
		}
	
	
	;Figure out what type of machine this is and write according file
	IfInString, The_SystemName, DDS
	{
	Fn_CheckDDS(The_SystemName)
	}
	
	IfInString, The_SystemName, IVR
	{
	Fn_CheckIVR(The_SystemName)
	}
	
	IfInString, The_SystemName, SVC
	{
	Fn_CheckSVC(The_SystemName)
	}
	
	IfInString, The_SystemName, SMW
	{
	Fn_CheckSMW(The_SystemName)
	}
	
	IfInString, The_SystemName, CBS
	{
	Fn_CheckDDS(The_SystemName)
	Fn_CheckTGP(The_SystemName)
	;Fn_CheckSVC(The_SystemName) Kindof annoying because it overwrites DDS stuffs
	}
	
	IfInString, The_SystemName, BOP
	{
	Fn_CheckBOP(The_SystemName)
	}
	
	IfInString, The_SystemName, MON
	{
	Fn_CheckMON(The_SystemName)
	}
	
	IfInString, The_SystemName, TGP
	{
	Fn_CheckTGP(The_SystemName)
	}
	
	IfInString, The_SystemName, MOB
	{
	;MOBs need separate handing as they do not have a registry
	Fn_CheckMOB(The_SystemName)
	}
}
;Clear the System in the GUI
guicontrol, Text, GUI_CurrentSystem, Sorting Excel Sheets...
Fn_GUI_UpdateProgress(10)

	;Loop every active connection file
	Loop, %Dir_ActiveConnections%\*.txt {
	The_SystemName := Fn_QuickRegEx(A_LoopFileName,"([a-zA-Z]*\d{2})")
		Loop, % Software_Array.MaxIndex() {
			If (Software_Array[A_Index,"SystemName"] = The_SystemName) {
			;Exit both Loops and do not move ACTIVE file to Archive
			Continue 2
			}
		}
	FileMove, %Dir_ActiveConnections%\%The_SystemName%.txt, %Dir_ArchiveConnections%\%The_SystemName%.txt
	}

;View the array
;Array_Gui(Software_Array)

;Create Excel Object Stuff
oExcel := ComObjCreate("Excel.Application") ; create Excel Application object
oExcel.Workbooks.Add ; create a new workbook (oWorkbook := oExcel.Workbooks.Add)
oExcel.Visible := False ; make Excel Application visible


;Read each system info file into Excel Page
;Parameter notes: Path to Active Connections, Platform, SystemType, OPTIONAL EXCEL PAGE OVERRIDE

	Fn_SortIntoExcel(Dir_ActiveConnections, "DSM", "IVR")
	Fn_SortIntoExcel(Dir_ActiveConnections, "PDX", "IVR")
	Fn_SortIntoExcel(Dir_ActiveConnections, "NJ", "IVR")
	Fn_SortIntoExcel(Dir_ActiveConnections, "VIA", "IVR")
	Fn_GUI_UpdateProgress(20)
	
	Fn_SortIntoExcel(Dir_ActiveConnections, "NJ", "SVC")
	Fn_SortIntoExcel(Dir_ActiveConnections, "VIA", "SVC")
	Fn_GUI_UpdateProgress(40)
	
	Fn_SortIntoExcel(Dir_ActiveConnections, "VIA", "BOP", "VIA_Misc")
	Fn_SortIntoExcel(Dir_ActiveConnections, "NJ", "BOP", "VIA_Misc")
	Fn_SortIntoExcel(Dir_ActiveConnections, "PHL", "BOP", "VIA_Misc")
	Fn_SortIntoExcel(Dir_ActiveConnections, "WOG", "UTILITY", "VIA_Misc")
	Fn_GUI_UpdateProgress(50)
	
	Fn_SortIntoExcel(Dir_ActiveConnections, "VIA", "SMW", "SMW")
	Fn_SortIntoExcel(Dir_ActiveConnections, "NJ", "SMW", "SMW")
	Fn_SortIntoExcel(Dir_ActiveConnections, "PDX", "SMW", "SMW")

	Fn_GUI_UpdateProgress(70)
	
	Fn_SortIntoExcel(Dir_ActiveConnections, "VIA", "MON", "MON")
	Fn_SortIntoExcel(Dir_ActiveConnections, "NJA", "MON", "MON")
	Fn_GUI_UpdateProgress(80)
	
	;MOBs do not connect to TGPs apparently
	;Fn_SortIntoExcel(Dir_ActiveConnections, "VIA", "MOB")
	;Fn_SortIntoExcel(Dir_ActiveConnections, "NJ", "MOB")
	
	;Fn_SortIntoExcel(Dir_ActiveConnections, "PA", "CBS") ;CBS machines no longer exist
	Fn_SortIntoExcel(Dir_ActiveConnections, "NJ", "DDS")
	Fn_SortIntoExcel(Dir_ActiveConnections, "TVG", "DDS")
	
Fn_GUI_UpdateProgress(90)
	
;Make Main sheets------------------------------------------------------------------------------------------------------
;Read all tgp connections in remote dir
	Loop, %Dir_ActiveConnections%\*.txt
	{
		;Read the contents of each file. At the top here we are only getting the system name
		Loop, Read, %A_LoopFileFullPath%
		{
		regexmatch(A_LoopReadLine, "Name = (\D+\d\d)", RE_Name_System)
			;IGNORE ALL THIS, UNNEEDED
			If (RE_Name_System1 != "")
			{
			System_Name = %RE_Name_System1%
			}
		;Get the primary TGP Connection out of file contents
		RegExMatch(A_LoopReadLine, "TGPChoice = (\D+\d\d)", RE_TGP_Number)
			If (RE_TGP_Number1 != "")
			{
			TGP_Number = %RE_TGP_Number1%
			}
		}
		
		;Skip systems we don't care about
		If (InStr(System_Name, "Utility") || InStr(System_Name, "UNI") || InStr(System_Name, "TMS") || InStr(System_Name, "99") || InStr(System_Name, "PDXOPSDDS") || InStr(System_Name, "CLS"))
		{
		;Skip rest of loop, essentially just skipping the part where the file is written
		Continue
		}
		
	;Append the whole system name to a file called TVG_TGP05 or similar
	FileAppend, %System_Name%`n, %Path_Temp_TGPs%\%TGP_Number%.txt
	}

	
;Make Software Versions Sheet------------------------------------------------------------------------------------------
oExcel.Worksheets.Add
oExcel.ActiveSheet.Name := "Software Versions"

	oExcel.Range("B2").Value := "Production Systems Software Audit"
	oExcel.Range("B2:F2").Merge(True)
	oExcel.Range("B2").Font.Size := 28	
	
	;Center   oExcel.Range("A1").HorizontalAlignment := -4108, oExcel.Selection.VerticalAlignment := -4108
	oExcel.Range("G2").Value := "Conducted: " . G_Day
	oExcel.Range("G2:H2").Merge(True)
	oExcel.Range("G2").Font.Size := 10
	
	oExcel.Range("B3:H3").Interior.ColorIndex := 16
	oExcel.Range("3:3").RowHeight := 4.5
	
	
	oExcel.Columns("A:Z").ColumnWidth := 16	;Set all Columns A-Z to a width of 16
	oExcel.Columns("D").ColumnWidth := 26	;Custom Column Width
	oExcel.Columns("J").ColumnWidth := 26	;Custom Column Width
	oExcel.Columns("A").ColumnWidth := 2	;Custom THIN Column Width

Pointer_Excel1 = 2

;;Export Software Array to Excel
Loop, % Software_Array.MaxIndex()
{
X := A_Index
Pointer_ExcelA = B
SystemName := Software_Array[A_Index,"SystemName"]
System_Type := Fn_QuickRegEx(SystemName,"(\D{3})\d{2}")
	Loop, % Software_Array[X].MaxIndex()
	{
		;Add space and row headers of new system type
		If (System_Type != Last_SystemType || Last_SystemType = "")
		{
		Pointer_Excel1 += 3
		Last_SystemType := System_Type
		oExcel.Range(Pointer_ExcelA . Pointer_Excel1).Value := System_Type
		oExcel.Range(Pointer_ExcelA . Pointer_Excel1).Interior.ColorIndex := 50
		Hold_ExcelPointerA := Pointer_ExcelA ;Hold the Pointer_ExcelA and give back after loop. Kinda dumb I know
			;for each software name
			Loop, % Software_Array[X].MaxIndex()
			{
			Header_SoftwareName := Software_Array[X][A_Index]
			Pointer_ExcelA := Fn_IncrementExcelColumn(Pointer_ExcelA, 1)
			oExcel.Range(Pointer_ExcelA . Pointer_Excel1).Value := Header_SoftwareName
			oExcel.Range(Pointer_ExcelA . Pointer_Excel1).Interior.ColorIndex := 50 ;Color it green
			}
		Pointer_ExcelA := Hold_ExcelPointerA ;Give the Pointer_ExcelA back
		}
		
		;Check if this is a new system and increment the excel pointer if so
		If(SystemName != Last_SystemName || Last_SystemName = "")
		{
		Last_SystemName := SystemName
		Pointer_Excel1 += 1
		
		oExcel.Range(Pointer_ExcelA . Pointer_Excel1) := SystemName
		oExcel.Range(Pointer_ExcelA . Pointer_Excel1).Interior.ColorIndex := 15 ;Color it light gray
		}
		
		;Put the software version in Excel if the version is not blank
		SoftwareName := Software_Array[X][A_Index]
		If (SoftwareName != "")
		{
		SoftwareVersion := Software_Array[X][SoftwareName]
			If (SoftwareVersion != "")
			{
			Pointer_ExcelA := Fn_IncrementExcelColumn(Pointer_ExcelA, 1)
			oExcel.Range(Pointer_ExcelA . Pointer_Excel1).Value := SoftwareVersion
			oExcel.Range(Pointer_ExcelA . Pointer_Excel1).Interior.ColorIndex := 35 ;Color it light green
			oExcel.Range(Pointer_ExcelA . Pointer_Excel1).HorizontalAlignment := -4108, oExcel.Selection.VerticalAlignment := -4108
			}
		}
	}
}
;view the array
;Array_Gui(Software_Array)



;Create the last page of Excel Sheet-----------------------------------------------------------------------------------
oExcel.Worksheets.Add
oExcel.ActiveSheet.Name := "All TGPs"

;Top header
	oExcel.Range("A1:E1").Merge(True)
	oExcel.Range("A1").Value := "Generated by System Reports " . The_VersionName
	oExcel.Range("A1").Interior.ColorIndex := 15
	oExcel.Range("A1").HorizontalAlignment := -4108, oExcel.Selection.VerticalAlignment := -4108
	
	oExcel.Range("A2:E2").Merge(True)
	oExcel.Range("A2").Value := "All TGPs and Connected Systems"
	oExcel.Range("A2:I2").Interior.ColorIndex := 16
	oExcel.Range("A2").HorizontalAlignment := -4108, oExcel.Selection.VerticalAlignment := -4108
	
	oExcel.Range("F1:I1").Interior.ColorIndex := 15
	oExcel.Range("F1").Value := G_Time
	oExcel.Range("G1").Value := G_Day
	
Pointer_ExcelA = A
Pointer_Excel1 = 3

Loop, %A_ScriptDir%\Data\temp\TGP_Connections\*.txt
{
TotalFiles := A_Index
}

Loop, %A_ScriptDir%\Data\temp\TGP_Connections\*.txt
{
	StringTrimRight, Current_TGP, A_LoopFileName, 4
	oExcel.Range(Pointer_ExcelA . Pointer_Excel1).Value := Current_TGP
	oExcel.Range(Pointer_ExcelA . Pointer_Excel1).Interior.ColorIndex := 50
	Pointer_ExcelB := Fn_IncrementExcelColumn(Pointer_ExcelA,1)
	;Make column wider
	oExcel.Columns(Pointer_ExcelA).ColumnWidth := 16
	oExcel.Columns(Pointer_ExcelB).ColumnWidth := 4
	oExcel.Range(Pointer_ExcelB . Pointer_Excel1).Value := "Ping"
	
	;Ping system and color cell based on result
	The_PingResult := Fn_Ping(System_Name)
		If (The_PingResult = 1)
		{
		oExcel.Range(Pointer_ExcelB . Pointer_Excel1).Interior.ColorIndex := 4
		}
		Else
		{
		oExcel.Range(Pointer_ExcelB . Pointer_Excel1).Interior.ColorIndex := 3
		}
		
	Loop, Read, %A_LoopFileFullPath%
	{
	;Spit all connected Systems under the TGP
	Pointer_Excel1 += 1
	oExcel.Range(Pointer_ExcelA . Pointer_Excel1).Value := A_LoopReadLine
	;Color B Cell according to result of earlier Ping test
	oExcel.Range(Pointer_ExcelB . Pointer_Excel1).Interior.ColorIndex := Fn_PingResults(A_LoopReadLine)
	}
	
	;Add two blank lines between each TGP written to excel
	Pointer_Excel1 += 3
	If (Pointer_Excel1 > 40)
	{
	;Increment column by turning into ascii and adding two. Convert back to character
	Pointer_Excel1 = 3
	;Convert Pointer_ExcelA to a character code from its existing ASCII counterpart
	Pointer_ExcelA := Asc(Pointer_ExcelA)
	Pointer_ExcelA += 3
		If (Pointer_ExcelA > 122)
		{
		Msgbox, I'm sorry columns greater than Z are not able to be handled. The program will Exit.
		ExitApp
		}
	;Return Pointer_ExcelA as an ASCII Character
	Pointer_ExcelA := Chr(Pointer_ExcelA)
	}
	Fn_GUI_UpdateProgress(A_Index, TotalFiles)
}

;Progress Bar at 100% now
Fn_GUI_UpdateProgress(100)
Fn_GUI_UpdateProgress(1,1)

;Read discrepant_servers to right of main sheet
Path_Discrepant = %A_ScriptDir%\Data\temp\discrepant_servers
Pointer_ExcelA := Asc(Pointer_ExcelA)
Pointer_ExcelA := Chr(Pointer_ExcelA+2)
Fn_SortIntoExcel(Path_Discrepant, "", "", "All TGPs", Pointer_ExcelA)



;Wrap Up

;Show Excel if it was hidden
;oExcel.Visible := true ; make Excel Application visible


;Create Archive directory and get new path to current day
The_ArchivePath = %A_ScriptDir%\Reports\Archive
The_ArchivePath := CreateArchiveDirandFilename(The_ArchivePath)

;Set Exitcode to 107191159 for 07-19 11:59 generated off the CurrentTime; a global variable

FormatTime, ExitCode, The_CurrentTime, MMddHHmm
ExitCode := "1" . ExitCode

;Tell Excel it is saved so it won't bother with save confirm messages
oExcel.ActiveWorkbook.saved := true
	;Save Excel to path of today (also delete the same Minute file if it already exists)
	IfExist, %The_ArchivePath%.*
	{
	FileDelete, %The_ArchivePath%.xlsx
	}
	
;Save Excel file to The_ArchivePath and copy to Most_Recent location
oExcel.ActiveWorkbook.SaveAs(The_ArchivePath)
FileCopy, %The_ArchivePath%.xlsx, %A_ScriptDir%\MostRecent_SystemReports.xlsx, 1

;Remove all temp files
FileRemoveDir, %Path_Temp% , 1

oExcel.Quit
ExitApp, %ExitCode%



;/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\
; Functions
;\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/--\--/

Startup()
{
#NoTrayIcon
#SingleInstance force
}


Fn_IncrementExcelColumn(para_Column,para_IncrementAmmount)
{
;Convert Column to a character code from its existing ASCII counterpart
l_Column := Asc(para_Column)
l_Column += %para_IncrementAmmount%
	If (l_Column > 122)
	{
	Msgbox, I'm sorry columns greater than Z are not able to be handled. The program will Exit.
	ExitApp
	}
Return Chr(l_Column)
}


Fn_Ping(para_IP)
{
;Ping the system and skip if there is no reply.

l_PingTime := 0
	Loop, 4
	{
	l_PingTime += 120
	runwait, %comspec% /c ping %para_IP% -n 1 -w %l_PingTime%,,UseErrorLevel Hide
		If (!ErrorLevel) {
			Return 1
		}
	}
	Return 0
}


Fn_CreateTempDir(para_Dir)
{
FileCreateDir, %para_Dir%
}


Fn_GetVersionNumber(para_Path)
{
FileGetVersion, l_OutputVar, %para_Path%
	If (l_OutputVar != "")
	{
	Return %l_OutputVar%
	}
Return "null"
}


Fn_GetModifedTime(para_Path)
{
FileGetTime, l_OutputVar, %para_Path%, M
	If (l_OutputVar > 19991111)
	{
	FormatTime, l_OutputVar, %l_OutputVar%, M/dd/yyyy
	Return %l_OutputVar%
	}
Return "null"
}


Fn_PingResults(para_SystemName)
{
Global The_Ping_Results_Array
l_ArraySize := The_Ping_Results_Array.MaxIndex()
X = 0
	Loop, %l_ArraySize%
	{
	X += 1
		If( para_SystemName = The_Ping_Results_Array[A_Index,1])
		{
			If( The_Ping_Results_Array[A_Index,2] = 1)
			{
			Return 4
			}
			Else If( The_Ping_Results_Array[A_Index,2] = 0)
			{
			Return 3
			}
		}
	}
Return 44
;Msgbox, the system %para_SystemName% was not found. Major Error
}


Fn_SortIntoExcel(para_filespath, para_platform, para_type, override_sheetname = 0, override_ExcelA = "A")
{
Global oExcel, Dir_MasterConnections

;Default Starting Position of Excel Sheet
Pointer_ExcelA = B
Pointer_Excel1 = 2
	If (override_ExcelA != "A")
	{
	Pointer_ExcelA := Asc(override_ExcelA)
	;Active
	ExcelC_ALabel := Chr(Pointer_ExcelA) ;Typically A
	ExcelC_BLabel := Chr(Pointer_ExcelA+1) ;Typically B
	;Master
	ExcelC_DLabel := Chr(Pointer_ExcelA+3) ;Typically D
	ExcelC_ELabel := Chr(Pointer_ExcelA+4) ;Typically E
	
	
	;Make header and stuff for this discrepant servers case
	oExcel.Columns(ExcelC_ALabel).ColumnWidth := 17
	oExcel.Columns(ExcelC_BLabel).ColumnWidth := 17
	oExcel.Columns(ExcelC_DLabel).ColumnWidth := 17
	oExcel.Columns(ExcelC_ELabel).ColumnWidth := 17
	oExcel.Range(ExcelC_ALabel . Pointer_Excel1 . ":" . ExcelC_BLabel . Pointer_Excel1).Merge(True)
	oExcel.Range(ExcelC_DLabel . Pointer_Excel1 . ":" . ExcelC_ELabel . Pointer_Excel1).Merge(True)
	oExcel.Range(ExcelC_ALabel . Pointer_Excel1).Value := "Active"
	oExcel.Range(ExcelC_ALabel . Pointer_Excel1).Interior.ColorIndex := 15
	oExcel.Range(ExcelC_ALabel . Pointer_Excel1).HorizontalAlignment := -4108, oExcel.Selection.VerticalAlignment := -4108
	oExcel.Range(ExcelC_DLabel . Pointer_Excel1).Value := "Master"
	oExcel.Range(ExcelC_DLabel . Pointer_Excel1).Interior.ColorIndex := 15
	oExcel.Range(ExcelC_DLabel . Pointer_Excel1).HorizontalAlignment := -4108, oExcel.Selection.VerticalAlignment := -4108
	
	;Top header
	oExcel.Range(ExcelC_ALabel . "1:" . ExcelC_ELabel . "1").Merge(True)
	oExcel.Range(ExcelC_ALabel . "1").Value := "Servers With Discrepancies"
	oExcel.Range(ExcelC_ALabel . "1").Interior.ColorIndex := 16
	oExcel.Range(ExcelC_DLabel . "1:" . ExcelC_ELabel . "1").HorizontalAlignment := -4108, oExcel.Selection.VerticalAlignment := -4108
	
	
	Pointer_Excel1 += 1
	}
	Else
	{
	;Active
	ExcelC_ALabel = A
	ExcelC_BLabel = B
	;Master
	ExcelC_DLabel = D
	ExcelC_ELabel = E
	}


	;Name Sheet
	If (override_sheetname = 0)
	{
	l_SheetName = %para_platform%_%para_type%
	}
	Else
	{
	l_SheetName = %override_sheetname%
	}

	
	;Add new sheet if the name does not match the active sheet (WARNING THIS WILL CRASH THE ENTIRE PROGRAM IF THE SHEETNAME EXISTS ELSEWHERE)
	If (oExcel.ActiveSheet.Name != l_SheetName)
	{
	oExcel.Worksheets.Add
	oExcel.ActiveSheet.Name := l_SheetName
	
	;All this junk just handles merging and stuff of the Active/Master Labels on Row 1
	oExcel.Columns("A").ColumnWidth := 17
	oExcel.Columns("B").ColumnWidth := 17
	oExcel.Columns("D").ColumnWidth := 17
	oExcel.Columns("E").ColumnWidth := 17
	oExcel.Range("A1:B1").Merge(True)
	oExcel.Range("D1:E1").Merge(True)
	oExcel.Range("A1").Value := "Active"
	oExcel.Range("A1").Interior.ColorIndex := 15
	oExcel.Range("A1").HorizontalAlignment := -4108, oExcel.Selection.VerticalAlignment := -4108
	oExcel.Range("D1").Value := "Master"
	oExcel.Range("D1").Interior.ColorIndex := 15
	oExcel.Range("D1").HorizontalAlignment := -4108, oExcel.Selection.VerticalAlignment := -4108
	}

	;Check if this starting position if blank; if not find a starting position
	While (oExcel.Range(ExcelC_ALabel . Pointer_Excel1).Value != "" || oExcel.Range(ExcelC_ALabel . Pointer_Excel1+1).Value != "" || oExcel.Range(ExcelC_ALabel . Pointer_Excel1+2).Value != "")
	{
	Pointer_Excel1 += 2
	}

	Loop, %para_filespath%\*.txt
	{
	TotalFiles := A_Index
	}
	
	;Read all files in the directory specified
	Loop, %para_filespath%\*.txt
	{
		;If they are part of the active selection (Example: NJ_SVC)
		If (InStr(A_LoopFileName, para_platform) && InStr(A_LoopFileName, para_type))
		{
			;Then read each line in the file
			Loop, Read, %A_LoopFileFullPath%
			{
			;Cut into Relevent Info and put into Excel
			StringSplit, Active_Array, A_LoopReadLine, `=,
			oExcel.Range(ExcelC_ALabel . Pointer_Excel1).Value := Active_Array1
			oExcel.Range(ExcelC_BLabel . Pointer_Excel1).Value := Active_Array2
			
			
			
			;Now input corresponding Master file, assuming it exists
			FileReadLine, ReadLine_Master, %Dir_MasterConnections%\%A_LoopFileName%, A_Index
			StringSplit, Master_Array, ReadLine_Master, `=,
			oExcel.Range(ExcelC_DLabel . Pointer_Excel1).Value := Master_Array1
			oExcel.Range(ExcelC_ELabel . Pointer_Excel1).Value := Master_Array2
			
			
			;Compare and highlight accordingly
				If (oExcel.Range(ExcelC_BLabel . Pointer_Excel1).Value = oExcel.Range(ExcelC_ELabel . Pointer_Excel1).Value)
				{
				oExcel.Range(ExcelC_BLabel . Pointer_Excel1).Interior.ColorIndex := 10
				}
				Else
				{
				oExcel.Range(ExcelC_BLabel . Pointer_Excel1).Interior.ColorIndex := 3
				
				;Lets remember this machine name since it doesn't match the master file
				FileCopy, %A_LoopFileFullPath%, %A_ScriptDir%\Data\temp\discrepant_servers\%A_LoopFileName%
				;FileAppend, "Mismatched master file", %A_ScriptDir%\Data\temp\discrepant_servers\%A_LoopFileName%
				}
			Pointer_Excel1 += 1
			}
		Pointer_Excel1 += 1
		Fn_GUI_UpdateProgress(A_Index, TotalFiles)
		}
	
	}

}


Fn_Txt2Array(Obj,para_TxtPath)
{
	Loop, Read, %para_TxtPath%
	{
	Obj.Insert(A_LoopReadLine)
	}
}

Fn_UpCase(para_Input) 
{
	StringUpper, OutputVar, para_Input
	Return OutputVar
}


Fn_ReadTGP(Remote_Machine)
{
RegRead, OutputVar, \\%Remote_Machine%\HKEY_LOCAL_MACHINE\SOFTWARE\TVG\SGRTransactionGateway, Machine
MachineName := Fn_QuickRegEx(OutputVar, "(\D+\d{2})")
Return Fn_UpCase(MachineName)
}


Fn_ReadSGRDataServ(Remote_Machine)
{
RegRead, OutputVar, \\%Remote_Machine%\HKEY_LOCAL_MACHINE\SOFTWARE\TVG\SGRDataService, Machine
MachineName := Fn_QuickRegEx(OutputVar, "(\D+\d{2})")
Return Fn_UpCase(MachineName)
}


Fn_ReadPrimaryDDS(Remote_Machine)
{
RegRead, OutputVar, \\%Remote_Machine%\HKEY_LOCAL_MACHINE\SOFTWARE\TVG\DdsControl, Machine
MachineName := Fn_QuickRegEx(OutputVar, "(\D+\d{2})")
Return Fn_UpCase(MachineName)
}


Fn_ReadBackupsDDS(Remote_Machine)
{
RegRead, Remote_Machine, \\%Remote_Machine%\HKEY_LOCAL_MACHINE\SOFTWARE\TVG\DdsControl, BackupMachine
MachineName := Fn_QuickRegEx(Remote_Machine, "(\D+\d{2})")
Return Fn_UpCase(MachineName)
}


Fn_ReadPrimaryDataCollector(Remote_Machine)
{
RegRead, Remote_Machine, \\%Remote_Machine%\HKEY_LOCAL_MACHINE\SOFTWARE\TVG\SGRDataService, SGRDataCollectorMachine
MachineName := Fn_QuickRegEx(Remote_Machine, "(\D+\d{2})")
Return Fn_UpCase(MachineName)
}


Fn_ReadBackupsDataCollector(Remote_Machine)
{
RegRead, Remote_Machine, \\%Remote_Machine%\HKEY_LOCAL_MACHINE\SOFTWARE\TVG\SGRDataService, SGRDataCollectorBackupMachine
MachineName := Fn_QuickRegEx(Remote_Machine, "(\D+\d{2})")
Return Fn_UpCase(MachineName)
}


Fn_ReadDatafeed(Remote_Machine)
{
RegRead, Remote_Machine, \\%Remote_Machine%\HKEY_LOCAL_MACHINE\SOFTWARE\TVG\SGRTransactionGateway, Machine
MachineName := Fn_QuickRegEx(Remote_Machine, "(\D+\d{2})")
Return Fn_UpCase(MachineName)
}


Fn_ReadVox(Remote_Machine)
{
RegRead, Remote_Machine, \\%Remote_Machine%\HKEY_LOCAL_MACHINE\SOFTWARE\TVG\IVR\STARTUP, IVRCallStatusServer
MachineName := Fn_QuickRegEx(Remote_Machine, "(\D+\d{2})")
Return Fn_UpCase(MachineName)
}

;NOT USED
Fn_ReadExperimental(l_RemoteMachine, l_key, l_data)
{
RegRead, l_RemoteMachine, \\%l_RemoteMachine%\HKEY_LOCAL_MACHINE\SOFTWARE\TVG\%l_key%, %l_data%
	regexmatch(1_Remote_Machine, "(\D+\d{2})", RE_RegData)
	If (RE_RegData1 != "")
	{
	l_FileredValue = %RE_RegData1%
	StringUpper, l_FileredValue, l_FileredValue
	}
Return l_FileredValue
}


Fn_CheckDDS(para_SystemName) {
l_TGP := Fn_ReadTGP(para_SystemName)
l_PrimaryDataCollector := Fn_ReadPrimaryDataCollector(para_SystemName)
l_BackupsDataCollector := Fn_ReadBackupsDataCollector(para_SystemName)

FileDelete, %A_ScriptDir%\Data\Active\%para_SystemName%.txt
FileAppend,
(
Computer Name = %para_SystemName%
TGPChoice = %l_TGP%
SDCChoice = %l_PrimaryDataCollector%
SDBChoice = %l_BackupsDataCollector%
), %A_ScriptDir%\Data\Active\%para_SystemName%.txt


Fn_CheckSoftwareVersion(para_SystemName,"OpCon.exe")
Fn_CheckSoftwareVersion(para_SystemName,"DDS Control.exe")
Fn_CheckSoftwareVersion(para_SystemName,"Raceday Explorer")
Fn_CheckSoftwareVersion(para_SystemName,"SGR Data Service.exe")
Fn_CheckSoftwareVersion(para_SystemName,"TMS")
Fn_CheckSoftwareVersion(para_SystemName,"DDS Web Server")
}


Fn_CheckIVR(para_SystemName) {
l_TGP := Fn_ReadTGP(para_SystemName)
l_PrimaryDDS := Fn_ReadPrimaryDDS(para_SystemName)
l_BackupsDDS := Fn_ReadBackupsDDS(para_SystemName)
l_Vox := Fn_ReadVox(para_SystemName)

FileDelete, %A_ScriptDir%\Data\Active\%para_SystemName%.txt
FileAppend,
(
Computer Name = %para_SystemName%
Primary DDS = %l_PrimaryDDS%
Backup DDS = %l_BackupsDDS%
TGPChoice = %l_TGP%
Vox Server = %l_Vox%
), %A_ScriptDir%\Data\Active\%para_SystemName%.txt

Fn_CheckSoftwareVersion(para_SystemName,"TMS")
Fn_CheckSoftwareVersion(para_SystemName,"IVR Service.exe")
Fn_CheckSoftwareVersion(para_SystemName,"DDS Web Server")
}


Fn_CheckMON(para_SystemName) {
l_TGP := Fn_ReadTGP(para_SystemName)
l_PrimaryDDS := Fn_ReadPrimaryDDS(para_SystemName)
l_BackupsDDS := Fn_ReadBackupsDDS(para_SystemName)

FileDelete, %A_ScriptDir%\Data\Active\%para_SystemName%.txt
FileAppend,
(
Computer Name = %para_SystemName%
TGPChoice = %l_TGP%
Primary DDS = %l_PrimaryDDS%
Backup DDS = %l_BackupsDDS%
), %A_ScriptDir%\Data\Active\%para_SystemName%.txt
}


Fn_CheckSVC(para_SystemName) {
l_TGP := Fn_ReadTGP(para_SystemName)
l_PrimaryDDS := Fn_ReadPrimaryDDS(para_SystemName)
l_BackupsDDS := Fn_ReadBackupsDDS(para_SystemName)

FileDelete, %A_ScriptDir%\Data\Active\%para_SystemName%.txt
FileAppend,
(
Computer Name = %para_SystemName%
Primary DDS = %l_PrimaryDDS%
Backup DDS = %l_BackupsDDS%
TGPChoice = %l_TGP%
), %A_ScriptDir%\Data\Active\%para_SystemName%.txt

Fn_CheckSoftwareVersion(para_SystemName,"TMS")
Fn_CheckSoftwareVersion(para_SystemName,"DDS Web Server")
Fn_CheckSoftwareVersion(para_SystemName,"Contest DDS Web Service")
Fn_CheckSoftwareVersion(para_SystemName,"Contest TMS Web Service")
Fn_CheckSoftwareVersion(para_SystemName,"DDS Web Service")
Fn_CheckSoftwareVersion(para_SystemName,"TMS Web Service")
}


Fn_CheckSMW(para_SystemName) {
l_TGP := Fn_ReadTGP(para_SystemName)
l_PrimaryDDS := Fn_ReadPrimaryDDS(para_SystemName)
l_BackupsDDS := Fn_ReadBackupsDDS(para_SystemName)

FileDelete, %A_ScriptDir%\Data\Active\%para_SystemName%.txt
FileAppend,
(
Computer Name = %para_SystemName%
Primary DDS = %l_PrimaryDDS%
Backup DDS = %l_BackupsDDS%
TGPChoice = %l_TGP%
), %A_ScriptDir%\Data\Active\%para_SystemName%.txt

Fn_CheckSoftwareVersion(para_SystemName,"TMS")
}


Fn_CheckTGP(para_SystemName) {
Fn_CheckSoftwareVersion(para_SystemName,"SGR Data Collector.exe")
Fn_CheckSoftwareVersion(para_SystemName,"SGR Transaction Gateway.exe")
}


Fn_CheckBOP(para_SystemName) {
l_TGP := Fn_ReadTGP(para_SystemName)
l_PrimaryDDS := Fn_ReadPrimaryDDS(para_SystemName)
l_BackupsDDS := Fn_ReadBackupsDDS(para_SystemName)
FileDelete, %A_ScriptDir%\Data\Active\%para_SystemName%.txt
FileAppend,
(
Computer Name = %para_SystemName%
Primary DDS = %l_PrimaryDDS%
Backup DDS = %l_BackupsDDS%
TGPChoice = %l_TGP%
), %A_ScriptDir%\Data\Active\%para_SystemName%.txt

Fn_CheckSoftwareVersion(para_SystemName,"Event Notification")
Fn_CheckSoftwareVersion(para_SystemName,"RaceDay Gateway")
}


Fn_CheckMOB(para_SystemName) {
Fn_CheckSoftwareVersion(para_SystemName,"DDS Support")
Fn_CheckSoftwareVersion(para_SystemName,"TVG HTTP Module")
Fn_CheckSoftwareVersion(para_SystemName,"TVG HTTP Module.pdb")
Fn_CheckSoftwareVersion(para_SystemName,"Xml Serializers")
Fn_CheckSoftwareVersion(para_SystemName,"TVG Mobile 2")
Fn_CheckSoftwareVersion(para_SystemName,"TVG Mobile 2.pdb")
Fn_CheckSoftwareVersion(para_SystemName,"TVG Mobile 2 Serializers")
Fn_CheckSoftwareVersion(para_SystemName,"TVG Shared")
Fn_CheckSoftwareVersion(para_SystemName,"TVG Shared.dll")
Fn_CheckSoftwareVersion(para_SystemName,"TVG Shared Serializers")
}


Fn_CheckWEB(para_SystemName) {
Fn_CheckSoftwareVersion(para_SystemName,"Error Logging Service")
Fn_CheckSoftwareVersion(para_SystemName,"Event Notification")
Fn_CheckSoftwareVersion(para_SystemName,"RaceDay Gateway")
Fn_CheckSoftwareVersion(para_SystemName,"Subscription Event Manager")
Fn_CheckSoftwareVersion(para_SystemName,"TVG Marketing Web Service")
}


Fn_CheckSoftwareVersion(para_SystemName,para_SoftwareName)
{ ;
Global

	Loop, % Software_Array.MaxIndex()
	{
		If (para_SystemName = Software_Array[A_Index, "SystemName"]) {
			X := A_Index
			Loop, % TxtDocument_Array.MaxIndex()
			{
				l_FileName := Fn_QuickRegEx(TxtDocument_Array[A_Index],"(\D+)#")
				l_FilePath := Fn_QuickRegEx(TxtDocument_Array[A_Index],"([a-z]\$.+)")
				l_FullPath = \\%para_SystemName%\%l_FilePath%

				If (l_FileName = para_SoftwareName) {
					l_versionnumber := ""
					l_versionnumber := Fn_GetVersionNumber(l_FullPath)
					If (l_versionnumber = "null" || l_versionnumber = "1.0.0.0" || l_versionnumber = "0.0.0.0") {
						l_versionnumber := Fn_GetModifedTime(l_FullPath)
					}
					If (l_versionnumber != "null") {
						Software_Array[X][l_FileName] := l_versionnumber
						Software_Array[X].Insert(l_FileName)
						Return
					}
				l_versionnumber := ""
				}
			}
		}
	}
}


CreateArchiveDirandFilename(para_StartDir)
{
global The_CurrentTime

;Returns currentday dir inside newly create 
The_CurrentTime := A_Now
FormatTime, l_CurrentYear, The_CurrentTime, yyyy
FormatTime, l_CurrentMonth, The_CurrentTime, MMMM
FormatTime, l_CurrentDay, The_CurrentTime, dd
FormatTime, l_CurrentTime, The_CurrentTime, HH-mm

FileCreateDir, %para_StartDir%\%l_CurrentYear%\%l_CurrentMonth%\%l_CurrentDay%\
l_FilePath = %para_StartDir%\%l_CurrentYear%\%l_CurrentMonth%\%l_CurrentDay%\%l_CurrentTime%
Return l_FilePath
}


;/--\--/--\--/--\--/--\
; GUI - Graphical User Interface
;\--/--\--/--\--/--\--/

GUI()
{
global
Gui +AlwaysOnTop

Gui, Font, s14 w70, Arial
Gui, Add, Text, x2 y0 w280 h40 +Center, System Reports %The_VersionName%
Gui, Font, s10 w70, Arial
Gui, Add, Text, x2 y20 w280 h20 +Center vGUI_CurrentSystem,

Gui, Show, x127 y87 h80 w288, System Reports

Gui, Add, Progress, x2 y50 w280 h10 -0x00000001 vGUI_ProgressBar1, 1
Gui, Add, Progress, x2 y60 w280 h20 -0x00000001 vGUI_ProgressBar2, 1

Return
}


Fn_GUI_UpdateProgress(Progress1, Progress2 = 0)
{

	If (Progress2 = 0)
	{
	;control big bar
	GuiControl,, GUI_ProgressBar2, %Progress1%+
	}
	Else
	{
	;control smaller bar
	Progress1 := (Progress1 / Progress2) * 100
	GuiControl,, GUI_ProgressBar1, %Progress1%
	}

}


GUI_DiableAllButtons()
{
GuiControl, disable, Run Report
}


GUI_EnableAllButtons()
{
GuiControl, enable, Run Report
}


GuiClose:
oExcel.Quit
ExitApp


;/--\--/--\--/--\--/--\
; Subroutines
;\--/--\--/--\--/--\--/

;NOT USED. MAY BE NEEDED IF PDX_OPERATOR EVER DIES
Sb_UpdateInfo(UpdateFile)
{
FileDelete, %UpdateFile%
File = 0
	While (File = 0)
	{
		IfExist, %UpdateFile%
		{
		File = 1
		}
		Sleep 1000
	}

}


Sb_ClearOutofProductionMachines()
{
global

	Loop, %Dir_ActiveConnections%\*.txt
	{
	ExistingActiveFile = %A_LoopFileName%	
	StringTrimRight, ExistingActiveFile, ExistingActiveFile, 3
	ArraySize := The_Ping_Results_Array.MaxIndex()
	X = 0
	SystemNoLongerActive = 0
		Loop, %ArraySize%
		{
		X += 1
			If( The_Ping_Results_Array[X,1] = ExistingActiveFile)
			{
			SystemNoLongerActive += 1
			}
			
		}
		If( SystemNoLongerActive > 0)
		{
		
		}
	}
	
}


Sb_ClearOldData()
{
global

FileRemoveDir, %Path_Temp%, 1

FileCreateDir, %Path_Temp_Discrepant%
FileCreateDir, %Path_Temp_TGPs%
}