'===========================================================
'| AD-UsersLoggedIn.vbs                                    |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 04/14/10                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| This script will report which users are logged into a   |
'| computer or report which users are logged into all      |
'| domain computers.                                       |
'|                                                         |
'===========================================================
'|                                                         |
'| Usage:                                                  |
'|                                                         |
'| Run from command line:                                  | 
'| cscript AD-UsersLoggedIn.vbs [COMPNAME] > Results.csv   |
'| or                                                      |
'| cscript AD-UsersloggedIn.vbs > Results.csv              |
'|                                                         |
'===========================================================

'==============================================
'| Check if a "Computer Name" cmd line        |
'| variable was passed to the script          |
'==============================================
On error resume next
strComputer=WScript.Arguments.Item(0)
On error goto 0
If strComputer="" Then
	' No specific computer was specified, proceed to query all computers in domain
Else
	strPingStatus = PingStatus(strComputer)
		If strPingStatus = "Success" Then
			QPO ' Run Query Process Owner function
		Else
			WScript.Echo strComputer & ",Failed ping with: " & strPingStatus
		End If
	WScript.quit    
End If

'==============================================
'|   Enumerate All Computers Accounts in AD   |
'==============================================
Const ADS_SCOPE_SUBTREE = 2
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCOmmand.ActiveConnection = objConnection
objCommand.CommandText = _
	"Select Name, Location from 'LDAP://DC=exchange,DC=local' " _
	& "Where objectClass='computer'"  
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
	strComputer = objRecordSet.Fields("Name").Value
	objRecordSet.MoveNext
	'Check if computer account is obsolite
	If obsoliteness(strComputer) =0 Then
		'check computer is on and echo "logged in user" or "ping failure status"
		strPingStatus = PingStatus(strComputer)
		If strPingStatus = "Success" Then
			QPO 'Run Query Process Owner function
		Else
			WScript.Echo strComputer & ",Failed ping with: " & strPingStatus &","&time
		End If
	Else
		WScript.Echo strComputer & ",Identified as an obsolite machine account,"&time
	End If
Loop

WScript.Quit(0)

'==============================================
'|                 Functions                  |
'==============================================

'==============================================
'|            Obsoliteness function           |
'==============================================
Function obsoliteness(var)
	Set myRegExp = New RegExp
	myRegExp.IgnoreCase = True
	myRegExp.Pattern = "(^XC00)|(^RC00)|(^QC00)|(^PC00)|(^OC00)|(^JC00)|(^HC00)|(^FC00)|(^EC00)|(^DC00)|(^CC)|(^C00)"
	obsoliteness = myRegExp.test(var)
End Function

'==============================================
'|        Query Process Owner function        |
'==============================================
Function QPO 
	On error resume next   
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" _
		& strComputer & "\root\cimv2")
	Set colProcessList = objWMIService.ExecQuery _
		("Select * from Win32_Process Where Name = 'explorer.exe'")
	For Each objProcess in colProcessList
		objProcess.GetOwner strUserName, strUserDomain
		Wscript.Echo strComputer &",Is logged into by "&strUserDomain & "\" & strUserName &","&time
	Next
End Function

'==============================================
'|            Ping Status function            |
'==============================================
Function PingStatus(strComputer)
	On Error Resume Next
	strWorkstation = "."
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strWorkstation & "\root\cimv2")
	Set colPings = objWMIService.ExecQuery _
		("SELECT * FROM Win32_PingStatus WHERE Address = '" & strComputer & "'")
	For Each objPing in colPings
		Select Case objPing.StatusCode
			Case 0 PingStatus = "Success"
            Case 11001 PingStatus = "Status code 11001 - Buffer Too Small"
            Case 11002 PingStatus = "Status code 11002 - Destination Net Unreachable"
            Case 11003 PingStatus = "Status code 11003 - Destination Host Unreachable"
            Case 11004 PingStatus = _
				"Status code 11004 - Destination Protocol Unreachable"
            Case 11005 PingStatus = "Status code 11005 - Destination Port Unreachable"
            Case 11006 PingStatus = "Status code 11006 - No Resources"
            Case 11007 PingStatus = "Status code 11007 - Bad Option"
            Case 11008 PingStatus = "Status code 11008 - Hardware Error"
            Case 11009 PingStatus = "Status code 11009 - Packet Too Big"
            Case 11010 PingStatus = "Status code 11010 - Request Timed Out"
            Case 11011 PingStatus = "Status code 11011 - Bad Request"
            Case 11012 PingStatus = "Status code 11012 - Bad Route"
            Case 11013 PingStatus = "Status code 11013 - TimeToLive Expired Transit"
            Case 11014 PingStatus = _
				"Status code 11014 - TimeToLive Expired Reassembly"
            Case 11015 PingStatus = "Status code 11015 - Parameter Problem"
            Case 11016 PingStatus = "Status code 11016 - Source Quench"
            Case 11017 PingStatus = "Status code 11017 - Option Too Big"
            Case 11018 PingStatus = "Status code 11018 - Bad Destination"
            Case 11032 PingStatus = "Status code 11032 - Negotiating IPSEC"
            Case 11050 PingStatus = "Status code 11050 - General Failure"
            Case Else PingStatus = "Status code " & objPing.StatusCode & _
				" - Unable to determine cause of failure."
        End Select
    Next
	On error goto 0
End Function 