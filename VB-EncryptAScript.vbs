'===========================================================
'| VB-EncryptAScript.vbs                                   |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 04/10/10                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| Encrypt a VBscript with a password that must be entered |
'| to run the script.                                      |
'|                                                         |
'===========================================================
 
'Declare RC4 arrays
dim key(255)
dim sbox(255)

'MD5 initialise
Dim sDigest
Private m_lOnBits(30)
Private m_l2Power(30)

md5y=0
md5x=1
m_l2Power(md5y)=cLng(md5x)
do
	md5x=md5x*2
	m_lOnBits(md5y)=CLng(md5x-1)
	if md5y=30 then exit do
	md5y=md5y+1
	m_l2Power(md5y)=cLng(md5x)
loop

'Enter path to plain text script file
strPT=inputbox("Enter path to vbs file you would like to Wrap", "Script Wrapper")

'Enter key
psw=inputbox("Enter key", "RC4 Encryption", "Password")
pswMD5=MD5(psw)

'Read plain text script file
Set objFSO = CreateObject ("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(strPT)
strTxt = objFile.ReadAll
objFile.close

'Encrypt
strTemp = EnDeCrypt(strTxt, pswMD5)
for x = 1 to len(strTemp)
	strHexDump=strHexDump&right(string(2,"0") & hex(asc(mid(strTemp, x, 1))),2)
next

'Save as another file
objFSO.OpenTextFile(strPT&" RC4_encrypted_with "&psw&".vbs", 8, True).WriteLine "strHex="&chr(34)&strHexDump&chr(34)

'WrapperBody is the wrapper of the encrypted script. 
'All the <CR> and <"> characters have been replaced with some uncommon but printable characters 
'A convenient way of storing a block of code to be used as a template.  

wrapperBody="§¥§¥'MD5 initialise§¥Dim sDigest§¥Private m_lOnBits(30)§¥Private m_l2Power(30)§¥md5y=0 : md5x=1§¥m_l2Power(md5y)=cLng(md5x)§¥do§¥ md5x=md5x*2§¥ m_lOnBits(md5y)=CLng(md5x-1)§¥ if md5y=30 then exit do§¥ md5y=md5y+1§¥ m_l2Power(md5y)=cLng(md5x)§¥loop§¥§¥§¥'Declare RC4 arrays§¥dim key(255)§¥dim sbox(255)§¥§¥§¥'Check for licence key§¥RegHex=mid(strHex,16,32)§¥pswMD5=©©§¥pswMD5=Licence(RegHex,pswMD5)§¥§¥§¥'Enter licence key§¥if pswMD5=©© then§¥ psw=inputbox(©Enter licence key©,©RC4 Encryption©,©Password©)§¥ pswMD5=MD5(psw)§¥ licence RegHex,pswMD5§¥end if§¥§¥§¥'Decrypt§¥for x=1 to len(strHex) step 2§¥ cTxt=cTxt&chr(cint(©&H©&mid(strHex,x,2)))§¥next§¥strTemp=EnDeCrypt(cTxt,pswMD5)§¥§¥§¥'Run script§¥on error resume next§¥executeglobal strtemp§¥if err<>0 then §¥ if err.number=13 then 'Type Mismatch§¥  msgbox ©Incompatible Script Found©&vbcrlf&©Your Script has same variables as the wrapper.©&vbcrlf&©Or you tried to wrap an already wrapped script.©§¥ else§¥  msgbox ©Wrong Encryption Key©§¥ end if§¥ Licence RegHex,©Bad©§¥end if§¥§¥§¥§¥§¥§¥function xorHWHash(strHash)§¥'Generate Hardware specific unlockKey§¥strComputer=©.©§¥Set objWMIService = GetObject(©winmgmts:\\© & strComputer & ©\root\CIMV2©)§¥Set colSMBIOS = objWMIService.ExecQuery (©Select * from Win32_SystemEnclosure©) §¥For Each objSMBIOS in colSMBIOS §¥ strHW=strHW&objSMBIOS.SerialNumber§¥Next  §¥MD5HW=MD5(strHW)§¥for iXOR=1 to len(MD5HW) step 1§¥ iMD5HW=mid(MD5HW,iXOR,1):iHash=mid(strHash,iXOR,1)§¥ xorHWHash=lcase(xorHWHash & hex((©&H©&iMD5HW) xor (©&H©&iHash)))§¥next§¥end function§¥§¥§¥§¥§¥Function EnDeCrypt(plaintxt,psw)§¥'RC4 encryption/Decryption§¥dim temp,a,i,j,k,cipherby,cipher§¥i=0:j=0§¥RC4Initialize psw§¥For a=1 To Len(plaintxt)§¥ i=(i+1) Mod 256§¥ j=(j+sbox(i)) Mod 256§¥ temp=sbox(i)§¥ sbox(i)=sbox(j)§¥ sbox(j)=temp§¥ k=sbox((sbox(i)+sbox(j)) Mod 256)§¥ cipherby=Asc(Mid(plaintxt,a,1)) Xor k§¥ cipher=cipher&Chr(cipherby)§¥Next§¥EnDeCrypt=cipher§¥End Function§¥§¥Sub RC4Initialize(strPwd)§¥dim tempSwap,a,b§¥intLength=len(strPwd)§¥For a=0 To 255§¥ key(a)=asc(mid(strpwd,(a mod intLength)+1,1))§¥ sbox(a)=a§¥next§¥b=0§¥For a=0 To 255§¥ b=(b+sbox(a)+key(a)) Mod 256§¥ tempSwap=sbox(a)§¥ sbox(a)=sbox(b)§¥ sbox(b)=tempSwap§¥Next§¥End Sub§¥§¥§¥§¥§¥§¥Function Licence(Regkey,pswMD5)§¥'Registry read write for licence key§¥on error resume next§¥strComputer=©.©§¥Const HKEY_LOCAL_MACHINE=&H80000002§¥Set oReg=GetObject(©winmgmts:{impersonationLevel=impersonate}!\\© &_ §¥strComputer & ©\root\default:StdRegProv©)§¥strKeyPath=©Software\Microsoft\Scripts©§¥strValueName=Regkey§¥oReg.CreateKey HKEY_LOCAL_MACHINE,strKeyPath§¥if pswMD5=©© and pswMD5<>©Bad© then§¥ oReg.GetBinaryValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue§¥ For ibin=lBound(strValue) to uBound(strValue)§¥  if strValue(ibin) <> ©0© then§¥   hexKey=hexKey & ©%© & hex(strValue(ibin))§¥  end if§¥ Next§¥ Licence=xorHWHash(hexDecode(hexKey))§¥else§¥ if pswMD5=©Bad© then §¥  oReg.DeleteValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName§¥  exit function§¥ end if§¥ xorpswMD5=xorHWHash(pswMD5)§¥ LenpswMD5=len(xorpswMD5)§¥ For ibin=1 to LenpswMD5§¥  Binary=Binary & ©&H© & hex(ASC(mid(xorpswMD5,ibin,1))) & ©,&H00,©§¥ next§¥ Binary=Binary & ©&H00© & ©,&H00©§¥ arrayBinary=split(Binary,©,©)§¥ oReg.SetBinaryValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,arrayBinary§¥end if§¥on error goto 0§¥end function§¥§¥§¥§¥§¥§¥function hexDecode(str)§¥'HEX to String§¥dim strDecoded,i,hexValue§¥strDecoded=©©§¥for i=2 to Len(str)§¥ hexValue=©©§¥ while Mid(str,i,1) <> ©%© and i <= Len(str)§¥  hexValue=hexValue+Mid(str,i,1)§¥  i=i+1§¥ wend§¥ strDecoded=strDecoded+chr(CLng(©&h© & hexValue))§¥next  §¥hexDecode=strDecoded§¥end function§¥§¥§¥§¥§¥§¥Private Function LShift(lValue,iShiftBits)§¥'MD5 code from www.frez.co.uk§¥§¥If iShiftBits=0 Then§¥ LShift=lValue§¥ Exit Function§¥ElseIf iShiftBits=31 Then§¥ If lValue And 1 Then§¥  LShift=&H80000000§¥ Else§¥  LShift=0§¥ End If§¥ Exit Function§¥ ElseIf iShiftBits<0 Or iShiftBits>31 Then§¥  Err.Raise 6§¥End If§¥§¥If (lValue And m_l2Power(31-iShiftBits)) Then§¥ LShift=((lValue And m_lOnBits(31-(iShiftBits+1)))*m_l2Power(iShiftBits)) Or &H80000000§¥Else§¥ LShift=((lValue And m_lOnBits(31-iShiftBits))*m_l2Power(iShiftBits))§¥End If§¥End Function§¥§¥Private Function RShift(lValue,iShiftBits)§¥If iShiftBits=0 Then§¥ RShift=lValue§¥ Exit Function§¥ElseIf iShiftBits=31 Then§¥ If lValue And &H80000000 Then§¥  RShift=1§¥ Else§¥  RShift=0§¥ End If§¥ Exit Function§¥ElseIf iShiftBits<0 Or iShiftBits>31 Then§¥ Err.Raise 6§¥End If §¥RShift=(lValue And &H7FFFFFFE)\m_l2Power(iShiftBits)§¥If (lValue And &H80000000) Then§¥ RShift=(RShift Or (&H40000000 \ m_l2Power(iShiftBits-1)))§¥End If§¥End Function§¥§¥Private Function RotateLeft(lValue,iShiftBits)§¥ RotateLeft=LShift(lValue,iShiftBits) Or RShift(lValue,(32-iShiftBits))§¥End Function§¥§¥Private Function AddUnsigned(lX,lY)§¥ Dim lX4,lY4,lX8,lY8,lResult§¥ lX8=lX And &H80000000§¥ lY8=lY And &H80000000§¥ lX4=lX And &H40000000§¥ lY4=lY And &H40000000§¥ lResult=(lX And &H3FFFFFFF)+(lY And &H3FFFFFFF)§¥If lX4 And lY4 Then§¥ lResult=lResult Xor &H80000000 Xor lX8 Xor lY8§¥ElseIf lX4 Or lY4 Then§¥ If lResult And &H40000000 Then§¥  lResult=lResult Xor &HC0000000 Xor lX8 Xor lY8§¥ Else§¥  lResult=lResult Xor &H40000000 Xor lX8 Xor lY8§¥ End If§¥Else§¥  lResult=lResult Xor lX8 Xor lY8§¥End If§¥AddUnsigned=lResult§¥End Function§¥§¥Private Function F(x,y,z)§¥ F=(x And y) Or ((Not x) And z)§¥End Function§¥§¥Private Function G(x,y,z)§¥ G=(x And z) Or (y And (Not z))§¥End Function§¥§¥Private Function H(x,y,z)§¥ H=(x Xor y Xor z)§¥End Function§¥§¥Private Function I(x,y,z)§¥ I=(y Xor (x Or (Not z)))§¥End Function§¥§¥Private Sub FF(a,b,c,d,x,s,ac)§¥ a=AddUnsigned(a,AddUnsigned(AddUnsigned(F(b,c,d),x),ac))§¥ a=RotateLeft(a,s)§¥ a=AddUnsigned(a,b)§¥End Sub§¥§¥Private Sub GG(a,b,c,d,x,s,ac)§¥ a=AddUnsigned(a,AddUnsigned(AddUnsigned(G(b,c,d),x),ac))§¥ a=RotateLeft(a,s)§¥ a=AddUnsigned(a,b)§¥End Sub§¥§¥Private Sub HH(a,b,c,d,x,s,ac)§¥ a=AddUnsigned(a,AddUnsigned(AddUnsigned(H(b,c,d),x),ac))§¥ a=RotateLeft(a,s)§¥ a=AddUnsigned(a,b)§¥End Sub§¥§¥Private Sub II(a,b,c,d,x,s,ac)§¥ a=AddUnsigned(a,AddUnsigned(AddUnsigned(I(b,c,d),x),ac))§¥ a=RotateLeft(a,s)§¥ a=AddUnsigned(a,b)§¥End Sub§¥§¥Private Function ConvertToWordArray(sMessage)§¥ Dim lMessageLength,lNumberOfWords,lBytePosition,lByteCount,lWordCount§¥ Dim lWordArray() §¥ Const MODULUS_BITS=512§¥ Const CONGRUENT_BITS=448 §¥ lMessageLength=Len(sMessage) §¥ lNumberOfWords=(((lMessageLength+((MODULUS_BITS-CONGRUENT_BITS) \ 8)) \ (MODULUS_BITS \ 8))+1)*(MODULUS_BITS \ 32)§¥ ReDim lWordArray(lNumberOfWords-1) §¥ lBytePosition=0§¥ lByteCount=0§¥ Do Until lByteCount >= lMessageLength§¥  lWordCount=lByteCount \ 4§¥  lBytePosition=(lByteCount Mod 4)*8§¥  lWordArray(lWordCount)=lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage,lByteCount+1,1)),lBytePosition)§¥  lByteCount=lByteCount+1§¥ Loop§¥ lWordCount=lByteCount \ 4§¥ lBytePosition=(lByteCount Mod 4)*8§¥ lWordArray(lWordCount)=lWordArray(lWordCount) Or LShift(&H80,lBytePosition)§¥ lWordArray(lNumberOfWords-2)=LShift(lMessageLength,3)§¥ lWordArray(lNumberOfWords-1)=RShift(lMessageLength,29)§¥ ConvertToWordArray=lWordArray§¥End Function§¥§¥Private Function WordToHex(lValue)§¥ Dim lByte§¥ Dim lCount §¥ For lCount=0 To 3§¥  lByte=RShift(lValue,lCount*8) And m_lOnBits(8-1)§¥  WordToHex=WordToHex & Right(©0© & Hex(lByte),2)§¥ Next§¥End Function§¥§¥Public Function MD5(sMessage)§¥ Dim x,k,AA,BB,CC,DD,a,b,c,d §¥ Const S11=7§¥ Const S12=12§¥ Const S13=17§¥ Const S14=22§¥ Const S21=5§¥ Const S22=9§¥ Const S23=14§¥ Const S24=20§¥ Const S31=4§¥ Const S32=11§¥ Const S33=16§¥ Const S34=23§¥ Const S41=6§¥ Const S42=10§¥ Const S43=15§¥ Const S44=21§¥ x=ConvertToWordArray(sMessage) §¥ a=&H67452301§¥ b=&HEFCDAB89§¥ c=&H98BADCFE§¥ d=&H10325476§¥ For k=0 To UBound(x) Step 16§¥  AA=a:BB=b:CC=c:DD=d §¥  FF a,b,c,d,x(k+0),S11,&HD76AA478§¥  FF d,a,b,c,x(k+1),S12,&HE8C7B756§¥  FF c,d,a,b,x(k+2),S13,&H242070DB§¥  FF b,c,d,a,x(k+3),S14,&HC1BDCEEE§¥  FF a,b,c,d,x(k+4),S11,&HF57C0FAF§¥  FF d,a,b,c,x(k+5),S12,&H4787C62A§¥  FF c,d,a,b,x(k+6),S13,&HA8304613§¥  FF b,c,d,a,x(k+7),S14,&HFD469501§¥  FF a,b,c,d,x(k+8),S11,&H698098D8§¥  FF d,a,b,c,x(k+9),S12,&H8B44F7AF§¥  FF c,d,a,b,x(k+10),S13,&HFFFF5BB1§¥  FF b,c,d,a,x(k+11),S14,&H895CD7BE§¥  FF a,b,c,d,x(k+12),S11,&H6B901122§¥  FF d,a,b,c,x(k+13),S12,&HFD987193§¥  FF c,d,a,b,x(k+14),S13,&HA679438E§¥  FF b,c,d,a,x(k+15),S14,&H49B40821 §¥  GG a,b,c,d,x(k+1),S21,&HF61E2562§¥  GG d,a,b,c,x(k+6),S22,&HC040B340§¥  GG c,d,a,b,x(k+11),S23,&H265E5A51§¥  GG b,c,d,a,x(k+0),S24,&HE9B6C7AA§¥  GG a,b,c,d,x(k+5),S21,&HD62F105D§¥  GG d,a,b,c,x(k+10),S22,&H2441453§¥  GG c,d,a,b,x(k+15),S23,&HD8A1E681§¥  GG b,c,d,a,x(k+4),S24,&HE7D3FBC8§¥  GG a,b,c,d,x(k+9),S21,&H21E1CDE6§¥  GG d,a,b,c,x(k+14),S22,&HC33707D6§¥  GG c,d,a,b,x(k+3),S23,&HF4D50D87§¥  GG b,c,d,a,x(k+8),S24,&H455A14ED§¥  GG a,b,c,d,x(k+13),S21,&HA9E3E905§¥  GG d,a,b,c,x(k+2),S22,&HFCEFA3F8§¥  GG c,d,a,b,x(k+7),S23,&H676F02D9§¥  GG b,c,d,a,x(k+12),S24,&H8D2A4C8A§¥  HH a,b,c,d,x(k+5),S31,&HFFFA3942§¥  HH d,a,b,c,x(k+8),S32,&H8771F681§¥  HH c,d,a,b,x(k+11),S33,&H6D9D6122§¥  HH b,c,d,a,x(k+14),S34,&HFDE5380C§¥  HH a,b,c,d,x(k+1),S31,&HA4BEEA44§¥  HH d,a,b,c,x(k+4),S32,&H4BDECFA9§¥  HH c,d,a,b,x(k+7),S33,&HF6BB4B60§¥  HH b,c,d,a,x(k+10),S34,&HBEBFBC70§¥  HH a,b,c,d,x(k+13),S31,&H289B7EC6§¥  HH d,a,b,c,x(k+0),S32,&HEAA127FA§¥  HH c,d,a,b,x(k+3),S33,&HD4EF3085§¥  HH b,c,d,a,x(k+6),S34,&H4881D05§¥  HH a,b,c,d,x(k+9),S31,&HD9D4D039§¥  HH d,a,b,c,x(k+12),S32,&HE6DB99E5§¥  HH c,d,a,b,x(k+15),S33,&H1FA27CF8§¥  HH b,c,d,a,x(k+2),S34,&HC4AC5665 §¥  II a,b,c,d,x(k+0),S41,&HF4292244§¥  II d,a,b,c,x(k+7),S42,&H432AFF97§¥  II c,d,a,b,x(k+14),S43,&HAB9423A7§¥  II b,c,d,a,x(k+5),S44,&HFC93A039§¥  II a,b,c,d,x(k+12),S41,&H655B59C3§¥  II d,a,b,c,x(k+3),S42,&H8F0CCC92§¥  II c,d,a,b,x(k+10),S43,&HFFEFF47D§¥  II b,c,d,a,x(k+1),S44,&H85845DD1§¥  II a,b,c,d,x(k+8),S41,&H6FA87E4F§¥  II d,a,b,c,x(k+15),S42,&HFE2CE6E0§¥  II c,d,a,b,x(k+6),S43,&HA3014314§¥  II b,c,d,a,x(k+13),S44,&H4E0811A1§¥  II a,b,c,d,x(k+4),S41,&HF7537E82§¥  II d,a,b,c,x(k+11),S42,&HBD3AF235§¥  II c,d,a,b,x(k+2),S43,&H2AD7D2BB§¥  II b,c,d,a,x(k+9),S44,&HEB86D391 §¥  a=AddUnsigned(a,AA)§¥  b=AddUnsigned(b,BB)§¥  c=AddUnsigned(c,CC)§¥  d=AddUnsigned(d,DD)§¥ Next §¥ MD5=LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))§¥End Function"

'unpack
wrapperBody=replace(wrapperBody, chr(165), chr(10))
wrapperBody=replace(wrapperBody, chr(167), chr(13))
wrapperBody=replace(wrapperBody, chr(169), chr(34))
objFSO.OpenTextFile(strPT&" RC4_encrypted_with "&psw&".vbs", 8, True).WriteLine wrapperBody

'---
Sub RC4Initialize(strPwd)
	dim tempSwap
	dim a
	dim b
	intLength = len(strPwd)
	For a = 0 To 255
		key(a) = asc(mid(strpwd, (a mod intLength)+1, 1))
		sbox(a) = a
	next
	b = 0
	For a = 0 To 255
		b = (b + sbox(a) + key(a)) Mod 256
		tempSwap = sbox(a)
		sbox(a) = sbox(b)
		sbox(b) = tempSwap
	Next
End Sub

'---
Function EnDeCrypt(plaintxt, psw)
	dim temp
	dim a
	dim i
	dim j
	dim k
	dim cipherby
	dim cipher
	i = 0
	j = 0
	RC4Initialize psw
	For a = 1 To Len(plaintxt)
		i = (i + 1) Mod 256
		j = (j + sbox(i)) Mod 256
		temp = sbox(i)
		sbox(i) = sbox(j)
		sbox(j) = temp
		k = sbox((sbox(i) + sbox(j)) Mod 256)
		cipherby = Asc(Mid(plaintxt, a, 1)) Xor k
		cipher = cipher & Chr(cipherby)
	Next
	EnDeCrypt = cipher
End Function

'---
' ###################################################
' # Start MD5 Code written by http://www.frez.co.uk #
' ###################################################

Private Function LShift(lValue, iShiftBits)
	If iShiftBits = 0 Then
		LShift = lValue
		Exit Function
	ElseIf iShiftBits = 31 Then
		If lValue And 1 Then
			LShift = &H80000000
		Else
			LShift = 0
		End If
		Exit Function
	ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
		Err.Raise 6
	End If
	If (lValue And m_l2Power(31 - iShiftBits)) Then
		LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
	Else
		LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
	End If
End Function

Private Function RShift(lValue, iShiftBits)
	If iShiftBits = 0 Then
		RShift = lValue
		Exit Function
	ElseIf iShiftBits = 31 Then
		If lValue And &H80000000 Then
			RShift = 1
		Else
			RShift = 0
		End If
		Exit Function
	ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
		Err.Raise 6
	End If
	RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
	If (lValue And &H80000000) Then
		RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
	End If
End Function

Private Function RotateLeft(lValue, iShiftBits)
	RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function

Private Function AddUnsigned(lX, lY)
	Dim lX4
	Dim lY4
	Dim lX8
	Dim lY8
	Dim lResult
	lX8 = lX And &H80000000
	lY8 = lY And &H80000000
	lX4 = lX And &H40000000
	lY4 = lY And &H40000000
	lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
	If lX4 And lY4 Then
		lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
	ElseIf lX4 Or lY4 Then
		If lResult And &H40000000 Then
			lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
		Else
			lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
		End If
	Else
		lResult = lResult Xor lX8 Xor lY8
	End If
	AddUnsigned = lResult
End Function

Private Function F(x, y, z)
	F = (x And y) Or ((Not x) And z)
End Function

Private Function G(x, y, z)
	G = (x And z) Or (y And (Not z))
End Function

Private Function H(x, y, z)
	H = (x Xor y Xor z)
End Function

Private Function I(x, y, z)
	I = (y Xor (x Or (Not z)))
End Function

Private Sub FF(a, b, c, d, x, s, ac)
	a = AddUnsigned(a, AddUnsigned(AddUnsigned(F(b, c, d), x), ac))
	a = RotateLeft(a, s)
	a = AddUnsigned(a, b)
End Sub

Private Sub GG(a, b, c, d, x, s, ac)
	a = AddUnsigned(a, AddUnsigned(AddUnsigned(G(b, c, d), x), ac))
	a = RotateLeft(a, s)
	a = AddUnsigned(a, b)
End Sub

Private Sub HH(a, b, c, d, x, s, ac)
	a = AddUnsigned(a, AddUnsigned(AddUnsigned(H(b, c, d), x), ac))
	a = RotateLeft(a, s)
	a = AddUnsigned(a, b)
End Sub

Private Sub II(a, b, c, d, x, s, ac)
	a = AddUnsigned(a, AddUnsigned(AddUnsigned(I(b, c, d), x), ac))
	a = RotateLeft(a, s)
	a = AddUnsigned(a, b)
End Sub

Private Function ConvertToWordArray(sMessage)
	Dim lMessageLength
	Dim lNumberOfWords
	Dim lWordArray()
	Dim lBytePosition
	Dim lByteCount
	Dim lWordCount
	Const MODULUS_BITS = 512
	Const CONGRUENT_BITS = 448
	lMessageLength = Len(sMessage)
	lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ 8)) \ (MODULUS_BITS \ 8)) + 1) * (MODULUS_BITS \ 32)
	ReDim lWordArray(lNumberOfWords - 1)
	lBytePosition = 0
	lByteCount = 0
	Do Until lByteCount >= lMessageLength
		lWordCount = lByteCount \ 4
		lBytePosition = (lByteCount Mod 4) * 8
		lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
		lByteCount = lByteCount + 1
	Loop
	lWordCount = lByteCount \ 4
	lBytePosition = (lByteCount Mod 4) * 8
	lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)
	lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
	lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
	ConvertToWordArray = lWordArray
End Function

Private Function WordToHex(lValue)
	Dim lByte
	Dim lCount
	For lCount = 0 To 3
		lByte = RShift(lValue, lCount * 8) And m_lOnBits(8 - 1)
		WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
	Next
End Function

Public Function MD5(sMessage)
	Dim x
	Dim k
	Dim AA
	Dim BB
	Dim CC
	Dim DD
	Dim a
	Dim b
	Dim c
	Dim d
	Const S11 = 7
	Const S12 = 12
	Const S13 = 17
	Const S14 = 22
	Const S21 = 5
	Const S22 = 9
	Const S23 = 14
	Const S24 = 20
	Const S31 = 4
	Const S32 = 11
	Const S33 = 16
	Const S34 = 23
	Const S41 = 6
	Const S42 = 10
	Const S43 = 15
	Const S44 = 21
	x = ConvertToWordArray(sMessage)
	a = &H67452301
	b = &HEFCDAB89
	c = &H98BADCFE
	d = &H10325476
	For k = 0 To UBound(x) Step 16
		AA = a
		BB = b
		CC = c
		DD = d
		FF a, b, c, d, x(k + 0), S11, &HD76AA478
		FF d, a, b, c, x(k + 1), S12, &HE8C7B756
		FF c, d, a, b, x(k + 2), S13, &H242070DB
		FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
		FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
		FF d, a, b, c, x(k + 5), S12, &H4787C62A
		FF c, d, a, b, x(k + 6), S13, &HA8304613
		FF b, c, d, a, x(k + 7), S14, &HFD469501
		FF a, b, c, d, x(k + 8), S11, &H698098D8
		FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
		FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
		FF b, c, d, a, x(k + 11), S14, &H895CD7BE
		FF a, b, c, d, x(k + 12), S11, &H6B901122
		FF d, a, b, c, x(k + 13), S12, &HFD987193
		FF c, d, a, b, x(k + 14), S13, &HA679438E
		FF b, c, d, a, x(k + 15), S14, &H49B40821
		GG a, b, c, d, x(k + 1), S21, &HF61E2562
		GG d, a, b, c, x(k + 6), S22, &HC040B340
		GG c, d, a, b, x(k + 11), S23, &H265E5A51
		GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
		GG a, b, c, d, x(k + 5), S21, &HD62F105D
		GG d, a, b, c, x(k + 10), S22, &H2441453
		GG c, d, a, b, x(k + 15), S23, &HD8A1E681
		GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
		GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
		GG d, a, b, c, x(k + 14), S22, &HC33707D6
		GG c, d, a, b, x(k + 3), S23, &HF4D50D87
		GG b, c, d, a, x(k + 8), S24, &H455A14ED
		GG a, b, c, d, x(k + 13), S21, &HA9E3E905
		GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
		GG c, d, a, b, x(k + 7), S23, &H676F02D9
		GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
		HH a, b, c, d, x(k + 5), S31, &HFFFA3942
		HH d, a, b, c, x(k + 8), S32, &H8771F681
		HH c, d, a, b, x(k + 11), S33, &H6D9D6122
		HH b, c, d, a, x(k + 14), S34, &HFDE5380C
		HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
		HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
		HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
		HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
		HH a, b, c, d, x(k + 13), S31, &H289B7EC6
		HH d, a, b, c, x(k + 0), S32, &HEAA127FA
		HH c, d, a, b, x(k + 3), S33, &HD4EF3085
		HH b, c, d, a, x(k + 6), S34, &H4881D05
		HH a, b, c, d, x(k + 9), S31, &HD9D4D039
		HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
		HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
		HH b, c, d, a, x(k + 2), S34, &HC4AC5665
		II a, b, c, d, x(k + 0), S41, &HF4292244
		II d, a, b, c, x(k + 7), S42, &H432AFF97
		II c, d, a, b, x(k + 14), S43, &HAB9423A7
		II b, c, d, a, x(k + 5), S44, &HFC93A039
		II a, b, c, d, x(k + 12), S41, &H655B59C3
		II d, a, b, c, x(k + 3), S42, &H8F0CCC92
		II c, d, a, b, x(k + 10), S43, &HFFEFF47D
		II b, c, d, a, x(k + 1), S44, &H85845DD1
		II a, b, c, d, x(k + 8), S41, &H6FA87E4F
		II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
		II c, d, a, b, x(k + 6), S43, &HA3014314
		II b, c, d, a, x(k + 13), S44, &H4E0811A1
		II a, b, c, d, x(k + 4), S41, &HF7537E82
		II d, a, b, c, x(k + 11), S42, &HBD3AF235
		II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
		II b, c, d, a, x(k + 9), S44, &HEB86D391
		a = AddUnsigned(a, AA)
		b = AddUnsigned(b, BB)
		c = AddUnsigned(c, CC)
		d = AddUnsigned(d, DD)
	Next
	MD5 = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
End Function 