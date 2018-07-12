'===========================================================
'| VB-EncryptText.vbs                                      |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 04/14/10                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| Decrypt text created with encryption script.            |
'|                                                         |
'===========================================================

Set x = WScript.createobject("wscript.shell") 
txt = inputbox("Enter Text to be Decoded") 
msgbox Encode(txt) 
x.Run "%windir%\notepad"
Wscript.sleep 1000 
x.sendkeys Encode(txt)
Wscript.Quit(0)

Function encode(s) 
	For i = 1 To Len(s) 
		newtxt = Mid(s, i, 1) 
		newtxt = Chr(Asc(newtxt)-3) 
		coded = coded & newtxt 
	Next 
	Encode = coded 
End Function 