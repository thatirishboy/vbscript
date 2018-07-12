'===========================================================
'| HW-GetDellModel.vbs                                     |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 04/28/11                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| This script will return the Dell model number from a    |
'| list of computer names.                                 |
'|                                                         |
'===========================================================
'|                                                         |
'| Reqirements:                                            |
'|                                                         |
'| - Must be run as Domain Admin.                          |
'| - List of computers:  .\pcs.txt                         |
'|                                                         |
'===========================================================

strWinTitle = "Trey's Computer Model Script v1.0"

On error resume next
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject") 
Set objPCs = objFSO.OpenTextFile(".\pcs.txt")
Set objExplorer = CreateObject("InternetExplorer.Application")
	objExplorer.Navigate "about:blank"
	objExplorer.ToolBar = 0
	objExplorer.StatusBar = 0
	objExplorer.Width = 800
	objExplorer.Height = 600
	objExplorer.Left = 100
	objExplorer.Top = 100
	objExplorer.Visible = 1
'Do While (objExplorer.Busy)
'Loop

Set objDocument = objExplorer.Document
objDocument.Open
objDocument.Writeln "<html><body><table border=1>"

' Loop through computers in list
Do While NOT objPCs.AtEndOfStream
	strComputer = objPCs.ReadLine
	If Reachable(strComputer) Then
		Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
		Set colItems1 = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
		For Each objItem In colItems1
			model = objitem.model
		Next
		If err.number = 0 Then
			objDocument.Writeln "<tr><td>" & strComputer & "</td><td>" & model & "</td></tr>"
		Else
			objDocument.Writeln "<tr><td>" & strComputer & "</td><td> Error" & err.number & "</td></tr>"
		End If
	Else
		objDocument.Writeln "<tr><td>" & strComputer & "</td><td>Unreachable</td></tr>"
	End If
Loop
objDocument.Writeln "</table></body></html>"
WScript.Quit(0)

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