'===========================================================
'| Win-CheckForService.vbs              |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 04/16/11                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| This script will check if a service input by the user   |
'| exists on all computers in the domain.                  |
'|                                                         |
'===========================================================
'|                                                         |
'| Reqirements:                                            |
'|                                                         |
'| - Must be run as Domain Admin.                          |
'|                                                         |
'===========================================================

Const OverwriteExisting = True
Const ADS_SCOPE_SUBTREE = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSysInfo = CreateObject("ADSystemInfo")
strWinTitle = "Trey's Service Checker Script v1.0"
strDomain = "Exchange"
strService = inputbox("Please enter the name of the service to check for." & vbcrlf & vbcrlf & "Note: Service names are case-sensitive.", strWinTitle)
If strService = "" Then
	WScript.Quit 0
End If

strLogName = "servicescript.log"
strLogPath = "s:\trey\logs\" & strLogName
On Error Resume Next
Set objLogFile = objFSO.OpenTextFile(strLogPath, 2, True) '"
objLogFile.WriteLine("############ Start Service Check Script ###############")
objLogFile.WriteLine("Checking for service: " & strService)
objLogFile.WriteLine("Script started " & Now)
objLogFile.WriteLine("")
If not err.number = 0 Then
	msgbox "There was a problem opening the log file for writing." & chr(10) & _
		"Please check whether """ & strLogPath & """ is a valid file and can be openend for writing." & _
		chr(10) & chr(10) & "If you're not sure what to do, please contact Trey.",vbCritical, strWinTitle
	ret = msgbox("Cannot find or open log file.", 0, strWinTitle)
	WScript.quit(1001)
End If

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCommand.ActiveConnection = objConnection
objCommand.CommandText = _
    "Select Name, Location from 'LDAP://" & strDomain & "' " _
        & "Where objectCategory='computer'"
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

On Error Resume Next
Do Until objRecordSet.EOF
	strCN = objRecordSet.Fields("Name").Value
	isServiceInstalled strService, strCN
	objRecordSet.MoveNext
Loop
objLogFile.WriteLine("")
objLogFile.WriteLine("Script completed " & Now)
objLogFile.WriteLine("############# End Serivce Check Script ################")
WScript.Echo "Done"
WScript.Quit 0

'==============================================
'|  Method to check if a service is installed |
'==============================================
Public Function isServiceInstalled(ByVal svcName, sCName)
	If Reachable(sCName) Then
		Set objWMIService = GetObject("winmgmts:\\" & sCName & "\root\CIMV2")
		' Obtain an instance of the the class
		' using a key property value.
		boolInstalled = FALSE
		svcQry = "SELECT * from Win32_Service"
		Set objOutParams = objWMIService.ExecQuery(svcQry)
		For Each objSvc in objOutParams
			If objSvc.Name = svcName Then
				boolInstalled = TRUE
			End If
		Next
		If boolInstalled = FALSE Then
			objLogFile.WriteLine("Not installed on " & sCName)
		End If
	End If
End Function 

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