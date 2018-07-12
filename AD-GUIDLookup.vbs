'===========================================================
'| AD-GUIDLookup.vbs                                       |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 02/12/11                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| This script will prompt for a GUID and convert it to    |
'| readable test from the Active Directory.                |
'|                                                         |
'===========================================================

strWinTitle = "Trey's GUID Convert Script v1.0"
strGUID = InputBox("Enter the GUID to look up.", strWinTitle)
Set obj = GetObject("LDAP://<GUID=" & DoIt(strGUID) & ">")
ret = MsgBox("GUID lookup result: " & vbcrlf & obj.Get("displayname"), 0, strWinTitle)

Function DoIt(strGUID)
    Dim octetStr, tmpGUID
    For i = 0 To Len(strGUID)
        t = Mid(strGUID, i + 1, 1)
        Select Case t
            Case "{"
            Case "}"
            Case "-"
            Case Else
                tmpGUID = tmpGUID + t
        End Select
    Next
    octetStr = Mid(tmpGUID, 7, 2)             ' 0
    octetStr = octetStr + Mid(tmpGUID,  5, 2) ' 1
    octetStr = octetStr + Mid(tmpGUID,  3, 2) ' 2
    octetStr = octetStr + Mid(tmpGUID,  1, 2) ' 3
    octetStr = octetStr + Mid(tmpGUID, 11, 2) ' 4
    octetStr = octetStr + Mid(tmpGUID,  9, 2) ' 5
    octetStr = octetStr + Mid(tmpGUID, 15, 2) ' 6
    octetStr = octetStr + Mid(tmpGUID, 13, 2) ' 7
    octetStr = octetStr + Mid(tmpGUID, 17, Len(tmpGUID))
    DoIt = octetStr
End Function 