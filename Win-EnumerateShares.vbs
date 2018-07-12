'===========================================================
'| Win-EnumerateShares.vbs                                 |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 05/09/11                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| Enumerate all shares and their permissions of a given   |
'| workstation or server.                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| Usage: csript /nologo <computername>                    |
'|                                                         |
'===========================================================

Option explicit

Dim strComputer, objWMI, colItems, strDir, i, objItem, wmiSecurityDescriptor, _
	wmiAce, strACE

strComputer = WScript.Arguments(0)

Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMI.ExecQuery("Select * from win32_share")

For Each i In colItems
	strDir = i.path
	WScript.Echo "Share Name: " & i.name
	strDir = Replace(strDir,"\","\\")
	Set colItems = objWMI.ExecQuery("Select * from win32_logicalFileSecuritySetting WHERE Path='" & strDir & "'",,48)
	for each objItem in colItems
		If objItem.GetSecurityDescriptor(wmiSecurityDescriptor) Then
			WScript.Echo "GetSecurityDescriptor failed"
			DisplayFileSecurity = False
			WScript.Quit
		End If
		For each wmiAce in wmiSecurityDescriptor.DACL
			strACE = wmiAce.Trustee.Domain & "\" & wmiAce.Trustee.Name
			'If instr(strACE,".") then
			wscript.echo " " & strACE
			'end If
		Next
	Next
Next 

Wscript.Quit(0) 