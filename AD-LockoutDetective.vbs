'===========================================================
'| AD-LockoutDetective.vbs                                 |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 02/15/11                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| If a user is reporting repeated lockouts of their       |
'| account, this script will identify the workstation      |
'| that is causing this.  An account lockout can be caused |
'| by but not limited to, a rogue user or a disused        |
'| workstation left logged into after a subsequent         |
'| password change.  This script will locate all your      |
'| Domain Controllers and scan their security event logs.  |
'| Filtered output.csv file is created in same folder as   |
'| script, to be interpreted in Excel.  Some variables can |
'| be modified in script setup section, User name and      |
'| number of days to search back are entered at script run |
'| time.                                                   |
'|                                                         |
'===========================================================

Option Explicit 

Dim objFSO, objWMI, objItem, intNumberID, colLoggedEvents, intEventType, strLogType, _
	strUser, DateFilter, objRootDSE, strConfig, adoCommand, adoConnection, adoRecordset, _
	strBase, strFilter, strAttributes, strQuery, UTCTime, strOutLogFile, k, y, z, objDC, _
	objDate 

' Script Setup 
strUser = inputbox("Enter UserName to filter on", "Lockout Detective", "XXX") 
DateFilter = inputbox("Enter number of days back you wish to check", "Lockout Detective", "1")
UTCTime = StdToUTCTime(Now()-DateFilter) 
intNumberID = "529"          ' Event ID Number for login failure. 
'intNumberID = "539"          ' Event ID Number for account lockout. 
intEventType = 5 '5=failures 4=success 
strLogType = "Security" 
strOutLogFile = "output.csv"

' Write output log file header 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
WriteLogFile "Controller,Time,User,Domain,Computer,Message",strOutLogFile 

' Determine configuration context and DNS domain from RootDSE object.
Set objRootDSE = GetObject("LDAP://RootDSE")
strConfig = objRootDSE.Get("configurationNamingContext") 

' Identify all Domain Controllers. 
Set adoCommand = CreateObject("ADODB.Command") 
Set adoConnection = CreateObject("ADODB.Connection") 
adoConnection.Provider = "ADsDSOObject" 
adoConnection.Open "Active Directory Provider" 
adoCommand.ActiveConnection = adoConnection 
strBase = "<LDAP://" & strConfig & ">" 
strFilter = "(objectClass=nTDSDSA)" 
strAttributes = "AdsPath" 
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree" 
adoCommand.CommandText = strQuery 
adoCommand.Properties("Page Size") = 100 
adoCommand.Properties("Timeout") = 60 
adoCommand.Properties("Cache Results") = False 
Set adoRecordset = adoCommand.Execute 

' Save Domain Controller AdsPaths into dynamic array arrstrDCs. 
k = 0 
Do Until adoRecordset.EOF 
	Set objDC = GetObject(GetObject(adoRecordset.Fields("AdsPath").Value).Parent) 
	ReDim Preserve arrstrDCs(k) 
	arrstrDCs(k) = objDC.DNSHostName 
	k = k+1 
	adoRecordset.MoveNext 
Loop 

adoRecordset.Close 

' Feed EACH Domain Controller name to event log reader 
For k = 0 To Ubound(arrstrDCs) 
	if instr(arrstrDCs(k), "ROOT") = 0 then  'Skip root domain controllers. 
		'msgbox arrstrDCs(k)
		EvLFilter arrstrDCs(k), strLogType, intEventType, intNumberID, UTCTime, strOutLogFile
	end if 
Next 
msgbox "Done" 

' Clean up. 
adoConnection.Close 
Set objRootDSE = Nothing 
Set adoConnection = Nothing 
Set adoCommand = Nothing 
Set adoRecordset = Nothing 
Set objDC = Nothing 
set objFSO = Nothing  

'----------------------------------------------------------- 
sub EvLFilter(DC, strLogType, intEventType, intNumberID, UTCTime, strOutLogFile) 
' Loop Through Filtered Event Logs, writing to output log file.
	Set objWMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & DC & "\root\cimv2") 
	Set colLoggedEvents = objWMI.ExecQuery ("SELECT * FROM Win32_NTLogEvent WHERE LogFile ='" & strLogType &_ 
		"' AND EventCode = '" & intNumberID & "' AND TimeGenerated >= '" & UTCTime & "'") 
	For Each objItem in colLoggedEvents 
		If objItem.EventType = intEventType then 
			if ucase(ParseUser(objItem.Message)) = ucase(strUser) or strUser = "" then
				WriteLogFile DC&","&ParseTime(objItem.TimeGenerated)&","&ParseUser(objItem.Message)&","&_
					ParseDomain(objItem.Message)&","&ParseWorkstation(objItem.Message)&","&_
					ParseMessage(objItem.Message),strOutLogFile  
			end if 
		End If 
	Next 
end sub 

'----------------------------------------------------------- 
Function ParseWorkstation(xxx) 
' Extract workstation name from event log 
	xxx=replace(xxx, vbtab, "") 
	z=instr(xxx, "Workstation Name:") : y=len(xxx) 
	xxx=right(xxx, y-z-16) 
	z=instr(xxx, vbcrlf) 
	ParseWorkstation=left(xxx, z-1) 
end Function 

'----------------------------------------------------------- 
Function ParseUser(xxx) 
' Extract user name from event log 
	xxx=replace(xxx, vbtab, "") 
	z=instr(xxx, "User Name:") : y=len(xxx) 
	xxx=right(xxx, y-z-9) 
	z=instr(xxx, vbcrlf) 
	ParseUser=left(xxx, z-1) 
end Function 

'----------------------------------------------------------- 
Function ParseDomain(xxx) 
' Extract Domain name from event log 
	xxx=replace(xxx, vbtab, "") 
	z=instr(xxx, "Domain:") : y=len(xxx) 
	xxx=right(xxx, y-z-6) 
	z=instr(xxx, vbcrlf) 
	ParseDomain=left(xxx, z-1) 
end Function 

'----------------------------------------------------------- 
Function ParseMessage(xxx) 
' Extract Reason Message from event log 
	xxx=replace(xxx, vbtab, "") 
	z=instr(xxx, "Reason:") : y=len(xxx) 
	xxx=right(xxx, y-z-6) 
	z=instr(xxx, vbcrlf) 
	ParseMessage=left(xxx, z-1) 
end Function 

'----------------------------------------------------------- 
Function ParseTime(xxx) 
' Convert UTC time to something more readable but still sortable in Excel 
	ParseTime =  Left(xxx, 4) &"/"& Mid(xxx, 5, 2) &"/"& (Mid(xxx, 7, 2) &" "&_ 
		Mid (xxx, 9, 2) &":"& Mid(xxx, 11, 2) &":"& Mid(xxx,13, 2)) 
end Function 

'----------------------------------------------------------- 
Function StdToUTCTime(inputDT)
' Convert Standard time to UTC format
	Dim objTime : Set objTime = CreateObject("WbemScripting.SWbemDateTime")
	Dim UTCDateTime: UTCDateTime = objtime.SetVarDate(inputDT)
	StdToUTCTime = objTime
End Function

'----------------------------------------------------------- 
Sub WriteLogFile(Content, FileName) 
' Write the filtered results to a csv file for close examination 
	objFSO.OpenTextFile(FileName, 8, True).WriteLine Content 
End Sub 