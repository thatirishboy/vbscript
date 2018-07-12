'===========================================================
'| IE-AddTrustedSite.vbs                                   |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 02/03/09                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| This script will add a registry entry to add a          |
'| specified domain to Internet Explorer Trusted Sites.    |
'|                                                         |
'===========================================================

On error resume next
Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
Set objReg=GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\" _
    & "ZoneMap\Domains\cutrainingonline.com" ' Path to entry; last item is domain to add
objReg.CreateKey HKEY_LOCAL_MACHINE, strKeyPath
strValueName = "http"
dwValue = 2
objReg.SetDWORDValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, dwValue
WScript.Quit(0) 