'===========================================================
'| Net-RemoveSpecifiedLocalAdministrators.vbs              |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 01/19/11                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| This script will remove specified users from the Local  |
'| Administrators group on specified computer.             |
'|                                                         |
'===========================================================
'|                                                         |
'| Reqirements:                                            |
'|                                                         |
'| - Must be run as Domain Admin.                          |
'|                                                         |
'===========================================================

Dim objGroup, objWshNet
Dim strWinTitle, strVersion, strUser, strCN, strDomain, boolExit, ret

boolExit = "N"
Set objWshNet = CreateObject("WScript.Network" ) 
strVersion = "v1.0"
strWinTitle = "Trey's Local Admin Cleaner Script " & strVersion
strDomain = objWshNet.UserDomain 

On error resume next
While boolExit = "N"
	err.clear
	strUser = InputBox("Enter the username to remove:", strWinTitle)
	If strUser = "" Then
		Wscript.Quit
	End If
	strCN = InputBox("Enter the computer name or IP:", strWinTitle) 
	If strCN = "" Then
		Wscript.Quit
	End If
	Set objGroup = GetObject("WinNT://" & strCN & "/Administrators,group")
	Set objUser = GetObject("WinNT://" & strDomain & "/" & strUser & ",user")
	objGroup.Remove(objUser.ADsPath)	
	If err.number = 0 Then
		msgbox "User " & strUser & " successfully removed from " & UCase(strCN) ,vbInformation , strWinTitle
	Else
		msgbox "Error removing " & strUser & " on " & UCase(strCN) & ".  Error: " & err.number, vbCritical, strWinTitle
	End If
	ret = MsgBox("Remove another user?", 4, strWinTitle)
	If ret = vbYes Then
		boolExit = "N"
	Else
		boolExit = "Y"
	End If
WEnd

Wscript.Quit 