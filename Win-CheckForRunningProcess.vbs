'===========================================================
'| Win-CheckForRunningProcess.vbs                          |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 08/27/09                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| This script will check running processes on a remote    |
'| computer and return if a specified process is running.  |
'|                                                         |
'===========================================================
'|                                                         |
'| Reqirements:                                            |
'|                                                         |
'| - Must be run as Domain Admin.                          |
'|                                                         |
'===========================================================

Function IsProcessRunning( strServer, strProcess )
	Dim Process, strObject
	IsProcessRunning = False
	strObject   = "winmgmts://" & strServer
	For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )
	If UCase( Process.name ) = UCase( strProcess ) Then
		IsProcessRunning = True
			Exit Function
		End If
	Next
End Function

Dim strComputer, strProcess, strWinTitle

strWinTitle = "Trey's Remote Running Process Script v1.0"

Do
	strProcess = inputbox("Please enter the name of the process (for instance: explorer.exe)", strWinTitle)
Loop until strProcess <> ""
Do
	strComputer = inputbox("Please enter the computer name", strWinTitle)
Loop until strComputer <> ""
If( IsProcessRunning( strComputer, strProcess ) = True ) Then
	ret = msgbox("Process " & strProcess & " is running on computer " & strComputer, 0, strWinTitle)
Else
	ret = msgbox("Process " & strProcess & " is NOT running on computer " & strComputer, 0, strWinTitle)
End If 
WScript.Quit(0) 