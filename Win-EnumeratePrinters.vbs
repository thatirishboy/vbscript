'===========================================================
'| Win-EnumeratePrinters.vbs                               |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 03/26/15                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| This script will collect a list of all mapped printers  |
'| for the currently logged in user and save them to a     |
'| text file on the S:\Trey\ drive.                        |
'|                                                         |
'===========================================================

Const ForWriting = 2

Set objNetwork = CreateObject("Wscript.Network")

strName = objNetwork.UserName
strDomain = objNetwork.UserDomain
strUser = strDomain & "\" & strName

strText = strUser & vbCrLf

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colPrinters = objWMIService.ExecQuery _
    ("Select * From Win32_Printer Where Local = FALSE")

For Each objPrinter in colPrinters
    strText = strText & objPrinter.Name & vbCrLf
Next

Set objFSO = CreateObject("Scripting.FileSystemObject")

strFileName = "S:\Trey\" & strName & "_Printers.txt"
wscript.echo strFileName

Set objFile = objFSO.CreateTextFile _
    (strFileName, ForWriting, True)

objFile.Write strText

objFile.Close