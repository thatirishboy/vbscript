'===========================================================
'| HW-GetDellSystemInfo.vbs                                |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 07/20/09                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| This script will prompt for an IP and will return the   |
'| serial number (Dell Service Tag) alone with other       |
'| computer and network information.                       |
'|                                                         |
'===========================================================
'|                                                         |
'| Reqirements:                                            |
'|                                                         |
'| - Must be run as Domain Admin.                          |
'|                                                         |
'===========================================================
'|                                                         |
'| Modified 08/28/09 by Trey: Added CPU type and speed     |
'| function and Focus XP info.                             |
'|                                                         |
'===========================================================

strWinTitle = "Trey's Hardware Info Script v1.0"

On error resume next
Do While strcomputer = "" AND a < 2
	strcomputer = Inputbox ("Please enter IP address or Computer name","Remote Computer Information Display","IP is preferred search method")
	strcomputer = trim(strcomputer)
	a = a + 1
Loop
If NOT strComputer <> "" Then
	ret = msgbox("No computer name entered, ending script", 0, strWinTitle)
	wscript.quit
End If

On error resume next
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set wbemServices = GetObject( "winmgmts://" & strComputer )
Set wbemObjectSet = wbemServices.InstancesOf( "Win32_LogicalMemoryConfiguration" )
Set objFSO = CreateObject("Scripting.FileSystemObject")

If err.number <> 0 Then
	If err.number = -2147217405 Then
		ret = msgbox("You do not have sufficient access rights to this computer", 0, strWinTitle)
		wscript.quit
	Else
		ret = msgbox("Could not locate computer" &vbcrlf& "Please check IP Address/Computer Name and try again", 0, strWinTitle)
		wscript.quit
	End If
End If

Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
Set colItems1 = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
Set colItems2 = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration Where IPEnabled = True")
Set colitems3 = objWMIService.ExecQuery("SELECT * FROM Win32_computersystem")
Set colitems4 = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkLoginProfile")
Set colitems5 = objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive")

Set objExplorer = CreateObject("InternetExplorer.Application")
objExplorer.Navigate "about:blank"
objExplorer.ToolBar = 0
objExplorer.StatusBar = 0
objExplorer.Width = 800
objExplorer.Height = 600
objExplorer.Left = 100
objExplorer.Top = 100
objExplorer.Visible = 1

Do While (objExplorer.Busy)
Loop

Set objDocument = objExplorer.Document
objDocument.Open
objDocument.Writeln "<html><head><title>Computer Information</title></head>"
objDocument.Writeln "<body bgcolor='white'>"

'==============================================
'|              Computer Details              |
'==============================================
For Each objItem In colItems
	serial = objitem.serialnumber
	BIOS = objitem.Name
Next
For Each objItem In colItems1
	hostname = objitem.caption
	make = objitem.manufacturer
	model = objitem.model
Next

objDocument.Writeln "<FONT color='red' size=4>Computer Information For: " & Ucase(hostname) & "</FONT><BR><BR>"
objDocument.Writeln "Serial : " & Serial & "</FONT><BR>"
objDocument.Writeln "BIOS : " & BIOS & "</FONT><BR>"
objDocument.Writeln "Make : " & make & "</FONT><BR>"
objDocument.Writeln "Model : " & Model & "</FONT><BR>"

Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor",,48)

For Each objItem in colItems
    objDocument.Writeln "CPU Type and Speed : " & objItem.Name & "<br>"
Next

For Each objItem In colItems5
	objDocument.Writeln "Drive Size: " & Int(objItem.Size /(1073741824)) & " GB<br>"
Next

For Each wbemObject In wbemObjectSet
	strMsg = vbCrLf & "Physical Memory: " & Int( ( wbemObject.TotalPhysicalMemory + 1023 ) / 1024 ) & " MB" & "<br>"
	objDocument.Writeln strMsg
Next

If objFSO.FileExists("\\" & strComputer & "\c$\windows\focusxp.ini") Then
	objDocument.Writeln "Focus XP Workstation Number : " & ReadIni("\\" & strComputer & "\c$\windows\focusxp.ini", "xpwnav", "WS") & "<br>"
	objDocument.Writeln "Focus XP Printers (Report/Receipt/Check) : " & ReadIni("\\" & strComputer & "\c$\windows\focusxp.ini", "xpwnav", "Rptpr") & "/" & ReadIni("\\" & strComputer & "\c$\windows\focusxp.ini", "xpwnav", "Rctpr") & "/" & ReadIni("\\" & strComputer & "\c$\windows\focusxp.ini", "xpwnav", "Chkpr") & "<br>"
End If

objDocument.Writeln "<BR><FONT color='Blue' size=4>Please Wait, gathering more information...</FONT><BR><BR>"

'==============================================
'|                User Details                |
'==============================================
For Each objItem In colItems3
	loggedon = objitem.username
next
For Each objItem In colItems4
	cachedlog = objitem.name
	username = objitem.FullName
	passwordexpire = objitem.passwordexpires
	badpassword = objitem.badpasswordcount
	If loggedon = cachedlog Then
		objDocument.Writeln "<FONT color='red' size=4>User Information For...</FONT><BR>"
		objDocument.Writeln "<FONT color='red' size=4>" & username & "</FONT><BR><BR>"
		objDocument.Writeln "User Name :" & loggedon &"</FONT><BR>"
		objDocument.Writeln "Incorrect Password Attempts : " & badpassword &"</FONT><BR>"
		On error resume next
		Set objaccount = GetObject("WinNT://scit84.sagchip.org/" &objitem.caption & ",user")
		If Err.Number <> 0 Then
			objDocument.Writeln "unable to retrieve password expiration information</FONT><BR>"
		Else
			If objAccount.PasswordExpired = 1 Then
				objDocument.Writeln "<FONT face='courier' color='red'>Password has Expired!</FONT><BR>"
			Else
			objDocument.Writeln "Password Expires " & objAccount.PasswordExpirationDate & " </FONT><BR><BR>"
			End If
		End If
	End If
Next

'==============================================
'|           Network Adapter Details          |
'==============================================
For Each objItem In colItems2
	ipaddress = objitem.ipaddress(0)
	description = objitem.description
	DHCP = objitem.DHCPserver
	Domain = objitem.DNSdomain
	mac = objitem.MACaddress
	DNS = objitem.dnsserversearchorder(0)
	DNS1 = objitem.dnsserversearchorder(1)
	DNS2 = objitem.dnsserversearchorder(2)
	wins1 = objitem.winsprimaryserver
	wins2 = objitem.winssecondaryserver
	speed = objitem.Speed
	If NOT ipaddress = "0.0.0.0" Then
		objDocument.Writeln "<FONT color='red' size=4>Network Adapter Details For...</FONT><BR>"
		objDocument.Writeln "<FONT color='red' size=4>" & description & "</FONT><BR><BR>"
		objDocument.Writeln "IP Address :" & ipaddress &"</FONT><BR>"
		objDocument.Writeln "DHCP Server : " & DHCP &"</FONT><BR>"
		objDocument.Writeln "Domain Name : " & domain &"</FONT><BR>"
		objDocument.Writeln "MAC Address : " & mac &"</FONT><BR>"
		objDocument.Writeln "Primary DNS : " & DNS &"</FONT><BR>"
		objDocument.Writeln "Secondary DNS : " & DNS1 &"</FONT><BR>"
		objDocument.Writeln "Tertiary DNS : " & DNS2 &"</FONT><BR>"
		objDocument.Writeln "Primary WINS : " & wins1 &"</FONT><BR>"
		objDocument.Writeln "Secondary WINS : " & WINS2 &"</FONT><BR><BR>"
	End If
Next

objDocument.Writeln "<FONT color='Blue' size=4>Script Finished</FONT><BR><BR>" 
WScript.Quit(0)

Function ReadIni( myFilePath, mySection, myKey )
    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8
    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection
    Set objFSO = CreateObject( "Scripting.FileSystemObject" )
    ReadIni     = ""
    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )
    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )
            ' Check if section is found in the current line
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = Trim( objIniFile.ReadLine )
                ' Parse lines until the next section is reached
                Do While Left( strLine, 1 ) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                        ' Check if item is found in the current line
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            ReadIni = Trim( Mid( strLine, intEqualPos + 1 ) )
                            ' In case the item exists but value is blank
                            If ReadIni = "" Then
                                ReadIni = " "
                            End If
                            ' Abort loop when item is found
                            Exit Do
                        End If
                    End If
                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do
                    ' Continue with next line
                    strLine = Trim( objIniFile.ReadLine )
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
        objDocument.Writeln strFilePath & " does not exist.<br>"
    End If
End Function 