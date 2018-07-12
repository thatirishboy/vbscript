'===========================================================
'| AD-UnlockUser.vbs                                       |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 05/13/09                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| This script will prompt for a username to be unlocked   |
'| the Active Directory.                                   |
'|                                                         |
'===========================================================
'|                                                         |
'| Reqirements:                                            |
'|                                                         |
'| - Must be run as Domain Admin.                          |
'|                                                         |
'===========================================================

Dim strUsername, strDomain, strWinTitle

strWinTitle = "Trey's AD Unlock Script v1.0"
strUsername=InputBox("Enter the username to unlock.", strWinTitle)
strDomain = "exchange"

Set objUser = GetObject("WinNT://" & strDomain & "/" & strUsername)
If objUser.IsAccountLocked = True Then
   objUser.IsAccountLocked = False
   objUser.SetInfo
   ret = msgbox("Account unlocked", 0, strWinTitle)
Else
   ret = msgbox("Account not locked", 0, strWinTitle)
End if

Wscript.quit(0) 