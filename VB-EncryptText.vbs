'===========================================================
'| VB-EncryptText.vbs                                      |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 04/14/10                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| Encrypt text and open it in Notepad to be copied.  Use  |
'| decryption script to reverse the encryption.            |
'|                                                         |
'===========================================================

Set x = WScript.CreateObject("WScript.Shell")
txt = inputbox("Enter Text to be Encoded") 
msgbox Encode(txt) 
x.Run "%windir%\notepad"
Wscript.Sleep 1000 
x.sendkeys Encode(txt)
Wscript.Quit(0)

Function Encode(s) 
	For i = 1 To Len(s) 
		newtxt = Mid(s, i, 1) 
		newtxt = Chr(Asc(newtxt)+3) 
		coded = coded & newtxt 
	Next 
	Encode = coded 
End Function 