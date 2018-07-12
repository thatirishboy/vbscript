'===========================================================
'| HW-BIOSAudit.vbs                                        |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 02/10/11                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| This script will prompt for an computer name and return |
'| the model and BIOS version to a text file.              |
'===========================================================
'|                                                         |
'| Reqirements:                                            |
'|                                                         |
'| - Must be run as Domain Admin.                          |
'|                                                         |
'===========================================================

tScriptStart = Now()
strWinTitle = "Trey's BIOS Audit Script v1.0"
strDomain = "Exchange"
Const OverwriteExisting = True
Const ADS_SCOPE_SUBTREE = 2
REM strCN = InputBox("Enter the name of the computer to check: ", strWinTitle) 
REM If strCN = "" Then
	REM Wscript.Quit(0) 
REM End If

Set objFSO = CreateObject("Scripting.FileSystemObject")
If NOT objFSO.FolderExists("S:\Trey\Logs") Then
	objFSO.CreateFolder("S:\Trey\Logs")
End If
strLogName = "BIOSlog2.txt"
strLogPath = "S:\Trey\Logs\" & strLogName
Set objLogFile = objFSO.OpenTextFile(strLogPath, 2, True)
Set objPCs = objFSO.OpenTextFile(".\biospcs.txt")
objLogFile.WriteLine("Audit Run on " & tScriptStart)

Do While NOT objPCs.AtEndOfStream
	strCN = objPCs.ReadLine
	GetInfo strCN
Loop

tScriptEnd = TimeSpan(tScriptStart, Now)
objLogFile.WriteLine("=========================================" & vbCRLF & "Total run time: " & tScriptEnd)
ret = msgbox("All Done!", 0, strWinTitle)
WScript.Quit(0) 

Sub GetInfo(sCName)
	If Reachable(sCName) Then
		err.Clear
		Set objWMIService = GetObject("winmgmts:\\" & sCName & "\root\CIMV2")
		Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS",,48)
		Set colItems1 = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem",,48)
		Set colitems5 = objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive")
		Set wbemServices = GetObject( "winmgmts://" & sCName )
		Set wbemObjectSet = wbemServices.InstancesOf( "Win32_LogicalMemoryConfiguration" )
		For Each objItem In colItems
			strBIOS = objitem.Name
		Next
		For Each objItem In colItems1
			strModel = objitem.model
		Next
		objLogFile.WriteLine("=========================================")
		objLogFile.WriteLine("Computer: " & UCase(sCName) & vbCRLF & "Model: " & strModel & vbCRLF & "BIOS: " & strBIOS)
	Else
		objLogFile.WriteLine("Cannot connect to " & sCName)
	End If
End Sub

Function Reachable(strComputer)
	On Error Resume Next
	Dim wmiQuery, objWMIService, objPing, objStatus
	wmiQuery = "Select * From Win32_PingStatus Where Address = '" & strComputer & "'"
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set objPing = objWMIService.ExecQuery(wmiQuery)
	For Each objStatus in objPing
		If IsNull(objStatus.StatusCode) Or objStatus.Statuscode<>0 Then
			Reachable = False 'if computer is unreacable, return false
		Else
			Reachable = True 'if computer is reachable, return true
		End If
	Next
End Function

Function TimeSpan(dt1, dt2)
	If (isDate(dt1) AND IsDate(dt2)) = False Then
		TimeSpan = "00:00"
		Exit Function
    End If
    seconds = Abs(DateDiff("S", dt1, dt2))
    minutes = seconds \ 60
    minutes = minutes mod 60
	seconds = seconds mod 60
	TimeSpan = RIGHT("00" & minutes, 2) & ":" & RIGHT("00" & seconds, 2)
End Function 