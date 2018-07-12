'===========================================================
'| Net-RebootEmAll.vbs              |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 05/13/09                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| This script will force logoff and restart all computers |
'| on all domains.                                         |
'|                                                         |
'===========================================================
'|                                                         |
'| Reqirements:                                            |
'|                                                         |
'| - Must be run as Domain Admin.                          |
'|                                                         |
'===========================================================

On error resume next

Set objNet = CreateObject("wscript.network")
strCurrentPC = objNet.ComputerName

'==============================================
'|          Connect to any AD domain          |
'==============================================
Const ADS_SCOPE_SUBTREE = 2
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection
Set objDSE = GetObject("LDAP://RootDSE")
strDomain = objDSE.Get("DefaultNamingContext")
objCommand.CommandText = "SELECT Name, Location FROM 'LDAP://" & strDomain & "' " & "WHERE objectClass=’computer’"
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Timeout") = 30
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
objCommand.Properties("Cache Results") = False
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
	strPCName = objRecordSet.Fields("Name").Value
	If strPCName <> strCurrentPC Then
		Set colOS = getobject("winmgmts:{impersonationlevel=impersonate,(shutdown)}//" & strPCName).instancesof("win32_operatingsystem")
		For Each objOS in colOS
			objOS.win32shutdown(2 + 4)
		Next
	End If
	objRecordSet.MoveNext
Loop
WScript.Quit(0) 