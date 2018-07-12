'===========================================================
'| Win-CheckRegistryPrivelages.vbs                         |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 05/22/09                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| This script will check whether the current user has     |
'| read and/or write access to predefined Registry keys    |
'| and values.                                             |
'|                                                         |
'===========================================================

Const KEY_QUERY_VALUE = &H0001
Const KEY_SET_VALUE = &H0002
Const KEY_CREATE_SUB_KEY = &H0004
Const DELETE = &H00010000

Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
strWinTitle = "Trey's Registry Access Script v1.0"

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
strComputer & "\root\default:StdRegProv")

strKeyPath = "SYSTEM\CurrentControlSet" ' Registry path to key

oReg.CheckAccess HKEY_LOCAL_MACHINE, strKeyPath, KEY_QUERY_VALUE, _
    bHasAccessRight
If bHasAccessRight = True Then
    ret = msgbox("Have Query Value Access Rights on Key", 0, strWinTitle)
Else
    ret = msgbox("Do Not Have Query Value Access Rights on Key", 0, strWinTitle)
End If   

oReg.CheckAccess HKEY_LOCAL_MACHINE, strKeyPath, KEY_SET_VALUE, _
    bHasAccessRight
If bHasAccessRight = True Then
    ret = msgbox("Have Set Value Access Rights on Key", 0, strWinTitle)
Else
    ret = msgbox("Do Not Have Set Value Access Rights on Key", 0, strWinTitle)
End If   
 
oReg.CheckAccess HKEY_LOCAL_MACHINE, strKeyPath, KEY_CREATE_SUB_KEY, _
    bHasAccessRight
If bHasAccessRight = True Then
    ret = msgbox("Have Create SubKey Access Rights on Key", 0, strWinTitle)
Else
    ret = msgbox("Do Not Have Create SubKey Access Rights on Key", 0, strWinTitle)
End If

oReg.CheckAccess HKEY_LOCAL_MACHINE, strKeyPath, DELETE, bHasAccessRight
If bHasAccessRight = True Then
    ret = msgbox("Have Delete Access Rights on Key", 0, strWinTitle)
Else
    ret = msgbox("Do Not Have Delete Access Rights on Key", 0, strWinTitle)
End If

Wscript.Quit(0) 