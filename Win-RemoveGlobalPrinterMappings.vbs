'===========================================================
'| Win-RemoveGlobalPrinterMappings.vbs              |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 03/17/10                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| This script will go into the local Registry to remove   |
'| all globally mapped printers from all profiles.         |
'|                                                         |
'===========================================================
'|                                                         |
'| Reqirements:                                            |
'|                                                         |
'| - Must be run as Domain Admin.                          |
'|                                                         |
'===========================================================

const HKEY_USERS = &H80000003
const HKEY_LOCAL_MACHINE = &H80000002

strWinTitle = "Trey's Global Printer Remover v1.0"
strComputer = "."
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

'==============================================
'| Routine to remove printer connections for  |
'| each profile                               |
'==============================================
strKeyPath = ""
objReg.EnumKey HKEY_USERS, strKeyPath, arrSubKeys
strKeyPath = "\Printers\Connections"

For Each subkey In arrSubKeys
	On Error Resume Next
	strKeyPath2 = subkey
	objReg.EnumKey HKEY_USERS, strKeyPath2 & strKeyPath, arrSubKeys2
	For Each strsubkey In arrSubKeys2
		On Error Resume Next
		strPrinter = subkey & strKeyPath & "\" & strsubkey
		objReg.EnumValues HKEY_USERS, strPrinter, arrValueNames
		For Each strValue in arrValueNames
			objReg.DeleteValue HKEY_USERS, strPrinter, strValue
		Next
		objReg.DeleteKey HKEY_USERS, strPrinter
	Next
Next

'==============================================
'| Routine to remove printer devices for      |
'| each profile                               |
'==============================================
strKeyPath = ""
objReg.EnumKey HKEY_USERS, strKeyPath, arrSubKeys
strKeyPath = "Software\Microsoft\Windows NT\CurrentVersion\Devices"

For Each subkey In arrSubKeys
	On Error Resume Next
	strKeyPath2 = subkey & "\" & strKeyPath
	objReg.EnumValues HKEY_USERS, strKeyPath2, arrValueNames
		For Each strValue in arrValueNames
			On Error Resume Next
			objReg.DeleteValue HKEY_USERS, strKeyPath2, strValue
		Next
Next

'==============================================
'| Routine to remove printer ports for each   |
'| profile                                    |
'==============================================
strKeyPath = ""
objReg.EnumKey HKEY_USERS, strKeyPath, arrSubKeys
strKeyPath = "Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts"

For Each subkey In arrSubKeys
	On Error Resume Next
	strKeyPath2 = subkey & "\" & strKeyPath
	objReg.EnumValues HKEY_USERS, strKeyPath2, arrValueNames
		For Each strValue in arrValueNames
			On Error Resume Next
			objReg.DeleteValue HKEY_USERS, strKeyPath2, strValue
		Next
Next

'==============================================
'| The following routines remove printers     |
'| the local machine                          |
'==============================================
strKeyPath = "SYSTEM\ControlSet001\Control\Print\Connections"
objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys

For Each subkey In arrSubKeys
	On Error Resume Next
	strKeyPath2 = subkey
	strPrinter = strKeyPath & "\" & strKeyPath2
	objReg.EnumValues HKEY_LOCAL_MACHINE, strPrinter, arrValueNames
	For Each strValue in arrValueNames
		On Error Resume Next
		objReg.DeleteValue HKEY_LOCAL_MACHINE, strPrinter, strValue
	Next
	objReg.DeleteKey HKEY_LOCAL_MACHINE, strPrinter
Next

strKeyPath = "SYSTEM\ControlSet003\Control\Print\Connections"
objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys

For Each subkey In arrSubKeys
	On Error Resume Next
	strKeyPath2 = subkey
	strPrinter = strKeyPath & "\" & strKeyPath2
	objReg.EnumValues HKEY_LOCAL_MACHINE, strPrinter, arrValueNames
	For Each strValue in arrValueNames
		On Error Resume Next
		objReg.DeleteValue HKEY_LOCAL_MACHINE, strPrinter, strValue
	Next
	objReg.DeleteKey HKEY_LOCAL_MACHINE, strPrinter
Next

strKeyPath = "SYSTEM\CurrentControlSet\Control\Print\Connections"
objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys

For Each subkey In arrSubKeys
	On Error Resume Next
	strKeyPath2 = subkey
	strPrinter = strKeyPath & "\" & strKeyPath2
	objReg.EnumValues HKEY_LOCAL_MACHINE, strPrinter, arrValueNames
	For Each strValue in arrValueNames
		On Error Resume Next
		objReg.DeleteValue HKEY_LOCAL_MACHINE, strPrinter, strValue
	Next
	objReg.DeleteKey HKEY_LOCAL_MACHINE, strPrinter
Next

ret = msgbox("Done", 0, strWinTitle)
Wscript.Quit(0) 