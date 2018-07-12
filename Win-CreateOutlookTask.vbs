'===========================================================
'| Win-CreateOutlookTask.vbs              |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 03/04/10                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| This script will prompt for fields and create an        |
'| Outlook task based on the responses.                    |
'|                                                         |
'===========================================================
'|                                                         |
'| Modified 04/15/10 by Trey: made the Cancel buttons quit |
'| the script without creating an empty task.              |
'|                                                         |
'===========================================================

Const olTaskItem = 3 
Dim strDate, strWinTitle

Set objOutlook = CreateObject("Outlook.Application") 
Set objTask = objOutlook.CreateItem(olTaskItem)
strWinTitle = "Trey's Outlook Task Creator v1.0"
strDate = InputBox("Enter the due date:", strWinTitle, Date())
If strDate = "" Then
	WScript.Quit
End If
With objTask
	.Subject = InputBox("Enter Task Subject:", strWinTitle, "Incident - ")
	.Body = "" 
	.ReminderSet = False
	.DueDate = strDate
	.StartDate = strDate
	.Categories = "Help Desk"
End With
If objTask.Subject = "" Then
	WScript.Quit
End If
objTask.Save
WScript.Quit(0) 