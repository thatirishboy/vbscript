'===========================================================
'| Net-EnumerateRDPSessions.vbs                            |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 05/21/12                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| Lists and optionally resets a user's server sessions.   |
'|                                                         |
'===========================================================

Option Explicit     

Const ADS_CHASE_REFERRALS_ALWAYS = &H20
Const ForAppend = 8

Dim wshShell, retval, oConn, oCmd, oRS, strADSPath, strADOQuery, strDomainCN, _
	fso,logfile, appendout, strUser, strSessionID

Set wshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

'Get the default ADsPath for the domain to search. 
Dim root: Set root = GetObject("LDAP://rootDSE")
strADSPath = root.Get("defaultNamingContext")

If (Not IsCScript()) Then 		'If not CScript, re-run with cscript...
	dim quote
	quote=chr(34)
	WshShell.Run "CScript.exe " & quote & WScript.ScriptFullName & quote, 1, true
    WScript.Quit            	'...and stop running as WScript
End If

If InStr(1,MyOS,"Server",1) = 0 Then
	MsgBox "You must run this from server OS",vbExclamation + vbOKOnly,"Error"
	'WScript.Quit
End If

retval = MsgBox("This script will identify and optionally logoff disconnected sessions for a user on all of the servers " & _
	"in AD within a domain. Do you want to continue?",vbYesNo + vbQuestion,"Get List of all Servers")
	If retval = vbNo Then WScript.Quit

strADSPath = InputBox("Get server list from what domain","Domain",strADSPath)
If strADSPath = "" Then WScript.Quit

strUser = InputBox("Search for what username?","User Name",wshShell.ExpandEnvironmentStrings("%USERNAME%"))
If strUser = "" Then WScript.Quit

Dim message
message = 	"Do you want to:" & VbCrLf & _
			"1) Get list only" & VbCrLf & _
			"2) Reset disconnected sessions" & VbCrLf & _
			"3) Reset all sessions for user" & VbCrLf & _	
			"0) Quit"	

Dim iActionType
iActionType  = InputBox(message,"Choose Action",1)
iActionType = CDbl(iActionType)
If iActionType = 0 Then WScript.Quit

GetServerList
wshShell.Run "notepad.exe " & quote & logfile & quote

' =========== Functions and Subs ==========

Sub GetServerList()
	'--- Set up the connection ---
	Set oConn = CreateObject("ADODB.Connection")
	Set oCmd = CReateObject("ADODB.Command")
	oConn.Provider = "ADsDSOObject"
	oConn.Open "ADs Provider"
	Set oCmd.ActiveConnection = oConn
	oCmd.Properties("Page Size") = 50
	ocmd.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
	logfile = Replace(strADSPath,",","_")
	logfile = Replace(logfile,"DC=","")
	logfile = wshShell.ExpandEnvironmentStrings("%userprofile%") & "\desktop\" & strUser & " In " &  logfile & ".txt" 
	If fso.FileExists(logfile) Then fso.DeleteFile logfile,True
	set AppendOut = fso.OpenTextFile(logfile, ForAppend, True)
	strDomainCN = DomainCN(strADSPath)
	'--- Build the query string ---
	strADOQuery = "<LDAP://" & strDomainCN & "/" & strADSPath & ">;" & "(&(OperatingSystem=*Server*)(objectClass=computer))" &  ";" & _
	    "Name;subtree"
	oCmd.CommandText = strADOQuery
	'--- Execute the query for the object in the directory ---
	Set oRS = oCmd.Execute
	If oRS.EOF and oRS.Bof Then
		MsgBox  "No Servers AD entries found!",vbCritical + vbOKOnly,"Failed"
		appendout.WriteLine "Query Failed"
	Else
		While Not oRS.Eof
			SessionQuery oRS.Fields("Name")
			oRS.MoveNext
		Wend
	End If
	
	oRS.Close
	oConn.Close
End Sub 

Sub SessionQuery (strServer)
	WScript.Echo "Checking " & strServer
	dim objEx, data
	Set objEx = WshShell.Exec("QWinsta /server:" & strServer)
	'one line at a time
	While Not (objEx.StdOut.AtEndOfStream)
		data = objEx.StdOut.ReadLine
		If InStr(1,data,strUser,1) Then
			strSessionID = GetSession(data)
			If iactionType = 1 then 
				EchoAndLog strServer & ",found session for " & strServer
			Else
				Wscript.echo strServer & ",found session for " & strServer
			End If 
			'always logoff
			If iActionType = 3 Then ResetSession strServer, strSessionID
			'Logoff disconnected
			If iActionType =2 And InStr(1,data,"disc",1) Then 
				ResetSession strServer,strSessionID
			End If 
		End If 
	Wend	
End Sub 

Sub ResetSession(strServer, ID)
	Dim strCommand, oExec
	strCommand = "reset session " & id & " /server:"  & strServer
	Set oExec    = WshShell.Exec(strCommand)
	wscript.sleep 500
	'this is typically empty
	While Not (oExec.StdOut.AtEndOfStream)
		EchoAndLog oExec.StdOut.ReadLine
	Wend
	If oExec.ExitCode <> 0 Then
		EchoAndLog strServer & ",Problem resetting session " & ID & " on server " & strServer & ", Non-zero exit code, " & oExec.exitcode
	Else
		EchoAndLog strServer & ",Reset session " & ID & " on server " & strServer
	End If
End Sub

Function DomainCN(strPath)
	DomainCN = Replace(strPath,",",".")		
	DomainCN= Replace(DomainCN,"DC=","")
End Function 

Function MyOS()
	Dim oWMI,ColOS,ObjOS, OSver
	Set oWMI = GetObject("winmgmts:\\.\root\cimv2")
	Set ColOS = oWMI.ExecQuery("SELECT Caption, version FROM Win32_OperatingSystem")
	For Each ObjOS In ColOS
		MyOS = objOS.caption & Space(1) & objos.version
	Next
End Function 

Function GetSession(text)
 	text = strip(lcase(Text))
 	Dim tArray, i
 	tArray = Split(text,Space(1))
 	i = 0 	
 	While tArray(i) <> lCase(strUser)
	 	i = i +1
 	Wend
 	GetSession = tArray(i+1)
End Function

Function Strip(text)
	text = Replace(text,vbtab,Space(1))
	While InStr(text,Space(2)) > 0
		text = replace(text,Space(2),Space(1))
	Wend
	Strip = text
End Function

Sub EchoAndLog (message)
	'Echo output and write to log
	Wscript.Echo message
	AppendOut.WriteLine message
End Sub  

Function IsCScript()
    If (InStr(UCase(WScript.FullName), "CSCRIPT") <> 0) Then
        IsCScript = True
    Else
        IsCScript = False
    End If
End Function
