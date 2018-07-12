'===========================================================
'| VB-ProgressBarSnippet.vbs                               |
'===========================================================
'|                                                         |
'| Created by:   Trey Donovan                              |
'| Last Updated: 07/20/11                                  |
'|                                                         |
'===========================================================
'|                                                         |
'| This snippet creates a progress bar using IE that can   |
'| be used in other scripts that perform many or long      |
'| functions.                                              |
'|                                                         |
'===========================================================

'==============================================
'|           Set up Progress Screen           |
'==============================================
Dim strTitle, objIE0

strTitle= "Progress Bar Title" 
nScrW= createobject("htmlfile").parentWindow.screen.availWidth 
nScrHt= createobject("htmlfile").parentWindow.screen.availHeight 
showBar objIE0, strTitle 
wscript.sleep 50
objIE0.document.parentWindow.document.script.listop "<br />"

strInsert = "Starting scan...<br />"
objIE0.document.parentWindow.document.script.listop strInsert

objIE0.quit
objIE0 = nothing
Wscript.quit

Sub showBar(robjIE0, usTitle) 
	Set robjIE0= createobject("internetExplorer.application") 
	robjIE0.navigate("about:blank") 
	Do 
		WScript.Sleep 50 
	Loop Until robjIE0.readyState=4 
	With robjIE0 
		.fullScreen= false
		.toolbar = false
		.statusBar = false
		.addressBar = false
		.resizable= false
		.menubar = false
		.width= 540 
		.height= 280 
		.left= (nScrW -520) \2 
		.top= (nScrHt -280) \2 
		With .document 
			.focus()
			.writeLn ("<!doctype html public>") 
			.writeLn ("<html style=""border-style:outset;" _ 
				& "border-width:4px"" " _ 
				& "onKeyDown=""vbscript:SuppressKeys"" " _ 
				& "onHelp=""vbscript:SuppressIeFns"" " _ 
				& "onContextMenu=""vbscript:SuppressIeFns"">") 
			.writeLn  ("<head>") 
			.writeLn   ("<title>" & usTitle & "</title>") 
			.writeLn   ("<style type=""text/css"">") 
			.writeLn    ("body {background-color:#ece9d8;" _ 
				& "text-align:center;" _ 
				& "vertical-align:middle}") 
			.writeLn   ("</style>") 
			.writeLn   ("<script language=""vbscript"">") 
			.writeLn    ("function SuppressKeys ()") 
			.writeLn     ("select case window.event.keyCode") 
			.writeLn      ("case 112, 114, 116") 
			.writeLn      ("case else: if NOT " _ 
				& "cbool(window.event.ctrlKey) then " _ 
				& "exit function") 
			.writeLn     ("end select") 
			.writeLn     ("window.event.keyCode= 0") 
			.writeLn     ("window.event.cancelBubble= true") 
			.writeLn     ("window.event.returnValue= false") 
			.writeLn    ("end function") 
			.writeLn    ("function SuppressIeFns ()") 
			.writeLn     ("window.event.cancelBubble= true") 
			.writeLn     ("window.event.returnValue= false") 
			.writeLn    ("end function") 
			.writeLn    ("function ListOp (ustrInsert)") 
			.writeLn     ("window.insertfile.insertAdjacentHtml " _ 
				& """beforeBegin"", ustrInsert") 
			.writeLn     ("window.insertfile.scrollIntoView") 
			.writeLn    ("end function") 
			.writeLn   ("</script>") 
			.writeLn  ("</head>") 
			.writeLn  ("<body scroll=""no"">") 
			.writeLn   ("<table>") 
			.writeLn    ("<tr>") 
			.writeLn     ("<td style=""text-align:center;" _ 
				& "font-family:Arial;font-size:16pt;" _ 
				& "font-weight:bold"">") 
			.writeLn      ("Processing Script...Please Wait") 
			.writeLn     ("</td>") 
			.writeLn    ("</tr>") 
			.writeLn    ("<tr>") 
			.writeLn     ("</td>") 
			.writeLn    ("</tr>") 
			.writeLn    ("<tr>") 
			.writeLn     ("<td style=""padding-top:15px"">") 
			.writeLn      ("<div id=""progresslist"" " _ 
				& "style=""height:150px;width:460px;" _ 
				& "max-height:100%;max-width:100%;" _ 
				& "padding-left:10px;text-align:left;" _ 
				& "font-family:Arial;font-size:10pt;" _ 
				& "font-weight:bold;border-style:inset;" _ 
				& "border-width:thin;overflow:scroll"">") 
			.writeLn       ("<span id=""insertfile""></span>") 
			.writeLn      ("</div>") 
			.writeLn     ("</td>") 
			.writeLn    ("</tr>") 
			.writeLn    ("<tr>") 
			.writeLn     ("<td style=""padding-top:20px;" _ 
				& "width:400px;font-family:Arial;" _ 
				& "font-size:10pt;" _ 
				& "font-weight:bold"">") 
			.writeLn     ("</td>") 
			.writeLn    ("</tr>") 
			.writeLn   ("</table>") 
			.writeLn  ("</body>") 
			.writeLn ("</html>") 
		End With 
		.visible= true 
	End With 
	WScript.Sleep 100 
	createobject("wscript.shell").appActivate usTitle 
End Sub 