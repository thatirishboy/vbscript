'===========================================================
'| Win-WindowsUpdates.vbs                                  |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 01/26/16                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| This script will enable and start services and scan for |
'| Windows Updates and prompt to install them.             |
'|                                                         |
'===========================================================
'|                                                         |
'| Reqirements:                                            |
'|                                                         |
'| - Must have internet connection.                        |
'|                                                         |
'===========================================================

Dim strVersion, boolGotoStop
strVersion = "1.0"
boolGotoStop = False

WScript.Echo "Trey's Windows Updater Script v" & strVersion & vbCrLf
WScript.Echo "Enabling and starting Windows Update services..."

Set objShell = WScript.CreateObject ("WScript.Shell")
objShell.run "sc config bits start= auto"
objShell.run "sc config wuauserv start= auto"
objShell.run "net start bits"
objShell.run "net start wuauserv"

Set updateSession = CreateObject("Microsoft.Update.Session")
updateSession.ClientApplicationID = "MSDN Sample Script"

Set updateSearcher = updateSession.CreateUpdateSearcher()

WScript.Echo "Searching for updates..." & vbCRLF

Set searchResult = _
updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

WScript.Echo "List of applicable items on the machine:"

For I = 0 To searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(I)
    WScript.Echo I + 1 & "> " & update.Title
Next

If searchResult.Updates.Count = 0 Then
    WScript.Echo "There are no applicable updates."
    boolGotoStop = True
End If

If boolGotoStop = False Then
	WScript.Echo vbCRLF & "Creating collection of updates to download:"

	Set updatesToDownload = CreateObject("Microsoft.Update.UpdateColl")

	For I = 0 to searchResult.Updates.Count-1
		Set update = searchResult.Updates.Item(I)
		addThisUpdate = false
		If update.InstallationBehavior.CanRequestUserInput = true Then
			WScript.Echo I + 1 & "> skipping: " & update.Title & _
			" because it requires user input"
		Else
			If update.EulaAccepted = false Then
				WScript.Echo I + 1 & "> note: " & update.Title & _
				" has a license agreement that must be accepted:"
				WScript.Echo update.EulaText
				WScript.Echo "Do you accept this license agreement? (Y/N)"
				strInput = WScript.StdIn.Readline
				WScript.Echo 
				If (strInput = "Y" or strInput = "y") Then
					update.AcceptEula()
					addThisUpdate = true
				Else
					WScript.Echo I + 1 & "> skipping: " & update.Title & _
					" because the license agreement was declined"
				End If
			Else
				addThisUpdate = true
			End If
		End If
		If addThisUpdate = true Then
			WScript.Echo I + 1 & "> adding: " & update.Title 
			updatesToDownload.Add(update)
		End If
	Next
End If

If updatesToDownload.Count = 0 Then
    WScript.Echo "All applicable updates were skipped."
    boolGotoStop = True
End If
    
If boolGotoStop = False Then
	WScript.Echo vbCRLF & "Downloading updates..."

	Set downloader = updateSession.CreateUpdateDownloader() 
	downloader.Updates = updatesToDownload
	downloader.Download()

	Set updatesToInstall = CreateObject("Microsoft.Update.UpdateColl")

	rebootMayBeRequired = false

	WScript.Echo vbCRLF & "Successfully downloaded updates:"

	For I = 0 To searchResult.Updates.Count-1
		set update = searchResult.Updates.Item(I)
		If update.IsDownloaded = true Then
			WScript.Echo I + 1 & "> " & update.Title 
			updatesToInstall.Add(update) 
			If update.InstallationBehavior.RebootBehavior > 0 Then
				rebootMayBeRequired = true
			End If
		End If
	Next
End If

If boolGotoStop = False Then
	If updatesToInstall.Count = 0 Then
		WScript.Echo "No updates were successfully downloaded."
		boolGotoStop = True
	End If

	If rebootMayBeRequired = true Then
		WScript.Echo vbCRLF & "These updates may require a reboot."
	End If
End If

If boolGotoStop = False Then
	WScript.Echo  vbCRLF & "Would you like to install updates now? (Y/N)"
	strInput = WScript.StdIn.Readline
	WScript.Echo 

	If (strInput = "Y" or strInput = "y") Then
		WScript.Echo "Installing updates..."
		Set installer = updateSession.CreateUpdateInstaller()
		installer.Updates = updatesToInstall
		Set installationResult = installer.Install()
 
		'Output results of install
		WScript.Echo "Installation Result: " & _
		installationResult.ResultCode 
		WScript.Echo "Reboot Required: " & _ 
		installationResult.RebootRequired & vbCRLF 
		WScript.Echo "Listing of updates installed " & _
		"and individual installation results:" 

		For I = 0 to updatesToInstall.Count - 1
			WScript.Echo I + 1 & "> " & _
			updatesToInstall.Item(i).Title & _
			": " & installationResult.GetUpdateResult(i).ResultCode   
		Next
	End If
End If

If rebootMayBeRequired = True Then
	WScript.Echo "Proceed with reboot?"
	If (strInput = "Y" or strInput = "y") Then
		WScript.Echo "Restarting computer..."
		objShell.run "shutdown -r -t 0"
	End If
End If

WScript.Echo  vbCRLF & "Would you like to disable services now? (Y/N)"
strInput = WScript.StdIn.Readline
WScript.Echo

If (strInput = "Y" or strInput = "y") Then
    WScript.Echo "Stopping and diabling Windows Update services..."
	objShell.run "net stop bits"
	objShell.run "net stop wuauserv"	
	objShell.run "sc config bits start= disabled"
	objShell.run "sc config wuauserv start= disabled"
End If

WScript.Echo  vbCRLF & "Would you like to run disk cleanup now? (Y/N)"
strInput = WScript.StdIn.Readline
WScript.Echo

If (strInput = "Y" or strInput = "y") Then
    WScript.Echo "Running disk cleanup..."
	objShell.run "c:\windows\system32\cleanmgr.exe /d c:"
End If

WScript.Quit