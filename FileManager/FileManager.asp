<%@codepage=1253 Language=VBScript%>

<%


	Const appName = "FileManager"
	Const appVersion = "2.00"
	
%>

	<!-- #include file="config.asp" -->
	
<%

	Dim FSO, re
	Dim scriptName, wexId
	Dim wexMessage, wexRootPath, targetPath
	Dim encoding, codepage, charset
	Dim FieldName
	
		
	InitApp()

	' Actions in the popup windows
	Select Case Request.Form("command")
		Case "Edit"
			Editor()
		Case "View"
			Viewer()
		Case "FileDetails", "FolderDetails"
			Details()
		Case "Upload"
			Upload(false)
	End Select
	
	' Actions in the main window
	Select Case Request.Form("command")
		Case "NewFile", "NewFolder"
			CreateItem()
		Case "DeleteFile", "DeleteFolder", "DeleteSelected"
			DeleteItem()
		Case "Unzip"
			Unzip()
		case "ZipSelectedFiles"
			ZipSelectedFiles()
		Case "RenameFile", "RenameFolder"
			RenameItem()
		Case "OpenFolder"
			targetPath = WexMapPath(Request.Form("folder") & Request.Form("parameter"))
		Case "LevelUp"
			targetPath = WexMapPath(FSO.GetParentFolderName(Request.Form("folder")))
		Case "Logout"
			Logout()
	End Select

	List()

	DestroyApp()

' ------------------------------------------------------------

' - WebExplorer Lite Functions -------------------------------

	' Initializes some variables, creates instances of some objects and ensures security
	Sub InitApp()
		
		GetLanguage(Language)
		
		scriptName = Request.ServerVariables("SCRIPT_NAME")
		wexId = appName & appVersion & "-"

		if trim(Request.QueryString("Field")) <> "" then
		
			Session("FieldName") = Request.QueryString("Field")
		
		end if
		
		If Request.QueryString("precommand")="Download" Then Response.Buffer = false Else Response.Buffer = true
		
		If not Secure() Then 
			If Request.Form("popup")="true" or Request.QueryString("popup")="true" Then PopupRelogin() Else Login()
		End If
		
		Set FSO = server.CreateObject ("Scripting.FileSystemObject")
		Set re = new regexp

		wexRootPath = RealizePath(wexRoot)

		encoding = -2 'System default encoding

		' Commands with high priority
		' These commands require to be performed before any Request.Form statement
		Select Case Request.QueryString("precommand")
			Case "ProcessUpload"
				Upload(true)
			Case "Download"
				Download()
			Case "Encoding"
				If Request.QueryString("value")<>"" Then encoding = Int(Request.QueryString("value"))
				If encoding=-1 Then 'Unicode encoding
					codepage = Session.CodePage
					Session.CodePage = 65001
					Response.CharSet = "UTF-8"
				End If
		End Select
		
		targetPath = WexMapPath(Request.Form("folder"))
	End Sub
	
	' Frees the objects and ends the application
	Sub DestroyApp()
		If encoding=-1 Then Session.CodePage = codepage
		Set FSO = Nothing
		Set re = Nothing
		Response.End
	End Sub
	
	Sub GetLanguage(LangFile)
	'########################
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	ReDim aTmp(0)
	f=Server.Mappath("lang/" & langfile)
	If fso.fileexists(f) Then
		Set fr=fso.OpenTextFile(f,1,False,-1)
		aLines=Split(fr.readall,VbCrLf)
		fr.close
		For n = 0 To UBound(aLines)
			s=Trim(aLines(n))
			Pos=Instr(s,"=")
			If s<>"" AND Pos>1 AND Pos<10 AND left(s,1)<>"'" AND left(s,1)<>";" Then 
				If Instr(s,";") Then s=Left(s,Instr(pos,s,";",1)-1)
				If IsNumeric(Left(s,Pos-1)) Then
					i=Int(Left(s,Pos-1))
						If i>Hi Then
						Hi=i
						Redim Preserve aTmp(i)
					End If
					aTmp(i)=Trim(Mid(s,Pos+1))
				End If
			End If
		Next
		Session("Str")=aTmp
	End If	
	Set fso=Nothing
	End SUB	
	
	' Writes the html header
	Sub HtmlHeader(title)
%>
<!--	<%=appName%> v<%=appVersion%>	Copyright © 2000-2007 Paradox Interactive	http://www.paradoxinteractive.com-->

<html>
<head>
<title><%=title%></title>
<meta NAME="keywords" CONTENT="Paradox,Interactive, i-paradox.gr,info@i-paradox.gr, Paradox Interactive, Web, Development, Graphics, Post Production, TV Commercials, Mutlimedia, CD Roms, DVD Authoring, 3D Modeling, 3D Animations, Logos, Presentations, PhotoAlbums, Vhs2DVD, Special FXs">
<meta NAME="description" CONTENT="Internet - Multimedia - DVD Authoring - TV Commercials - VHS2DVD - PhotoAlbums - eCommerce - 3D Animation - Video Wall Presentations">
<meta NAME="rating" CONTENT="General">
<meta NAME="revisit-after" CONTENT="30 days">
<meta NAME="objecttype" CONTENT="Homepage">
<meta NAME="MS.LOCALE" CONTENT="EL">
<meta NAME="author" CONTENT="George Papaioannou">
<meta NAME="company" CONTENT="Paradox Interactive">
<meta HTTP-EQUIV="creation_date" CONTENT="Sun, 1 Jan, 2007">
<meta HTTP-EQUIV="copyright" CONTENT="Paradox Interactive - 2007">
<meta HTTP-EQUIV="generator" CONTENT="MSHTML">
<%HtmlStyle%>
<%HtmlJavaScript%>
</head>
<body bgColor="#f79200">
<%	
	End Sub
	
	' Writes the html footer
	Sub HtmlFooter()
%>	
</body>
</html>
<%
	End Sub

	' Writes the copyright message
	Sub HtmlCopyright()
	%>
		<table cellspacing="0" cellpadding="0" border="0" align="center"><tr><td>
			<a target="_blank" href="http://www.paradoxinteractive.gr" title="paradox interactive">Copyright © 2000-2007 Paradox Interactive</a>
		</td></tr></table>
	<%
	End Sub
	
	' Writes the stylesheet
	Sub HtmlStyle()
%>
<style>
BODY
{
    BACKGROUND-COLOR: #FFFFFF
}
TD
{
    FONT-WEIGHT: normal;
    FONT-SIZE: 10pt;
    COLOR: #1B3986;
    FONT-FAMILY: Verdana,Tahoma,Arial,Helvetica
}
.formClass
{
    BACKGROUND-COLOR: #FFFFFF;
    FONT-WEIGHT: normal;
    FONT-SIZE: 10pt;
    COLOR: black;
    FONT-FAMILY: Verdana,Tahoma,Arial,Helvetica
}
.lightRow {
	BACKGROUND-COLOR: #EFEFF7;
	COLOR : #1B3986;
}
.darkRow {
	BACKGROUND-COLOR: #DEDEEF;
	COLOR : #1B3986;
	
}
.titleRow {
	BACKGROUND-COLOR: #74687D;

}
.loginRow {
	border: black solid 1px;
	BACKGROUND-COLOR: DodgerBlue
}
.boldText
{
    FONT-WEIGHT: bold;   
    FONT-SIZE: 10pt;
    COLOR: white;
    FONT-FAMILY: Verdana,Tahoma,Arial,Helvetica;    
}
.blackText
{
    FONT-WEIGHT: bold;
    COLOR: black;    
    FONT-SIZE: 10pt;    
    FONT-FAMILY: Verdana,Tahoma,Arial,Helvetica;    
}
A
{
    FONT-WEIGHT: bold;
    COLOR: #1B3986;
    TEXT-DECORATION: none
}
A:hover
{
    COLOR: #f79200;
    TEXT-DECORATION: none
}
A:visited
{
    TEXT-DECORATION: none
    
}
A:active
{
    COLOR: #74687D;
    TEXT-DECORATION: none
}
</style>
<%
	End Sub
	
	' Writes the javascript code
	Sub HtmlJavaScript()
%>
<script language="javascript">

	function selectAll(t)
	{
		
		selValue = false;
		stateValue= 0;
		
		if (t[t.length-1].checked == true) 		
			selValue = true;
				
		for(i=6; i<t.length-1; i++)	
			t[i].checked=selValue; 
		
			
		

	} 
		
	function Command(cmd, param) {
		var str;
		var someWin;
		switch (cmd) {
			case "NewFile":
				str = prompt("<%=Session("str")(43)%>", "New File");
				if(!str) return;
				else if (!CheckName(str)) {alert("<%=Session("str")(51)%>"); return;}
				document.forms.formBuffer.parameter.value = str;
				break;
			case "NewFolder":
				str = prompt("<%=Session("str")(52)%>", "New Folder");
				if(!str) return;
				else if (!CheckName(str)) {alert("<%=Session("str")(53)%>"); return;}
				document.forms.formBuffer.parameter.value = str;
				break;
			case "Edit":
				str = document.forms.formBuffer.folder.value + param;
				someWin = openWin(cmd + str, "", 600, 440, false, false);
				someWin.focus();
				createPage(someWin,cmd,param);
				someWin = null;
				return;
				break;
			case "View":
				str = document.forms.formBuffer.folder.value + param;				
				someWin = openWin(cmd + str, "", 600, 440, false, true);
				someWin.focus();
				createPage(someWin,cmd,param);
				someWin = null;
				return;
				break;
			case "FileDetails":
			case "FolderDetails":
				str = document.forms.formBuffer.folder.value + param;
				someWin = openWin(cmd + str, "", 350, 220, false, false);
				someWin.focus();
				createPage(someWin,cmd,param);
				someWin = null;
				return;
				break;
			case "Upload":
				someWin = openWin(cmd, "", 400, 150, true, false);
				someWin.focus(); 
				createPage(someWin,cmd,param);
				someWin = null;
				return;
				break;
			case "DeleteFolder":
				if (!confirm('<%=Session("str")(24)%> "' + param + ' ?')) return;
				document.forms.formBuffer.parameter.value = param;
				break;
			case "DeleteFile":
				if (!confirm('<%=Session("str")(23)%> "' + param + '" ?')) return;
				document.forms.formBuffer.parameter.value = param;				
				break;
			case "DeleteSelected":
				if (!confirm('<%=Session("str")(29)%> ?')) return;
				document.forms.formBuffer.parameter.value = param;				
				break;
			case "Unzip":
				if (!confirm('<%=Session("str")(41)%> ?')) return;
				document.forms.formBuffer.parameter.value = param;				
				break;	
			case "ZipSelectedFiles":
				str = prompt("<%=Session("str")(43)%>", "target.zip");
				if (!str) return;				
				document.forms.formBuffer.parameter.value = str;				
				break;											
			case "RenameFile":
				str = prompt("<%=Session("str")(54)%>", param);
				if(!str) return;
				else if (!CheckName(str)) {alert("<%=Session("str")(51)%>"); return;}
				document.forms.formBuffer.parameter.value = param + "|" + str;
				break;
			case "RenameFolder":
				str = prompt("<%=Session("str")(55)%>", param);
				if(!str) return;
				else if (!CheckName(str)) {alert("<%=Session("str")(53)%>"); return;}
				document.forms.formBuffer.parameter.value = param + "|" + str;
				break;
			case "NoWebAccess":
				alert("<%=Session("str")(56)%>");
				return;
				break;
			case "Zip":
			
				break;
				
			case "Unzip":
			
				break;
			default:
				document.forms.formBuffer.parameter.value = param;
		}
		document.forms.formBuffer.target = "";
		document.forms.formBuffer.command.value = cmd
		document.forms.formBuffer.submit();	
		
	}
	
	function Check() {
		if (document.forms.formBuffer.pwd.value == "") {
			alert("You haven't entered the password!"); 
			return false;
		} else return true;
	}

	function openWin(winName, urlLoc, w, h, showStatus, isViewer) {
		l = (screen.availWidth - w)/2;
		t = (screen.availHeight - h)/2;
		features  = "toolbar=no";      // yes|no 
		features += ",location=no";    // yes|no 
		features += ",directories=no"; // yes|no 
		features += ",status=" + (showStatus?"yes":"no");  // yes|no 
		features += ",menubar=no";     // yes|no 
		features += ",scrollbars=" + (isViewer?"yes":"no");   // auto|yes|no 
		features += ",resizable=" + (isViewer?"yes":"no");   // yes|no 
		features += ",dependent";      // close the parent, close the popup, omit if you want otherwise 
		features += ",height=" + h;
		features += ",width=" + w;
		features += ",left=" + l;
		features += ",top=" + t;
		winName = winName.replace(/[^a-z]/gi,"_");
		return window.open(urlLoc,winName,features);
	} 
	
	function createPage (theWin, cmd, param){
		document.forms.formBuffer.target = theWin.name;
		document.forms.formBuffer.command.value = cmd;
		document.forms.formBuffer.parameter.value = param;
		document.forms.formBuffer.popup.value = "true";
		document.forms.formBuffer.submit();
		document.forms.formBuffer.popup.value = "false";
	}

	function EditorCommand (cmd) {
		switch (cmd) {
			case "Info":
				alert(document.forms.formBuffer.info.value.replace(/\|/gi,"\n"));
				break;
			case "Reload":
				document.forms.formBuffer.reset();
				break;
			case "Save":
				document.forms.formBuffer.action += "?precommand=Encoding&value=";
				document.forms.formBuffer.action += document.forms.formBuffer.encoding.options[document.forms.formBuffer.encoding.selectedIndex].value;
				document.forms.formBuffer.subcommand.value = "Save";
				document.forms.formBuffer.submit();
				break;
			case "SaveAs":
				var str, oldname;
				oldname = document.forms.formBuffer.parameter.value;
				str = prompt("Save the file as :", oldname);
				if (!str || str==oldname) return;
				document.forms.formBuffer.action += "?precommand=Encoding&value=";
				document.forms.formBuffer.action += document.forms.formBuffer.encoding.options[document.forms.formBuffer.encoding.selectedIndex].value;
				document.forms.formBuffer.parameter.value = str;
				document.forms.formBuffer.subcommand.value = "SaveAs";
				document.forms.formBuffer.submit();
				break;
			case "Encoding":
				document.forms.formBuffer.action += "?precommand=Encoding&value=";
				document.forms.formBuffer.action += document.forms.formBuffer.encoding.options[document.forms.formBuffer.encoding.selectedIndex].value;
				document.forms.formBuffer.subcommand.value = cmd;
				document.forms.formBuffer.submit();
				break;
		}
	}

	function ViewerCommand (cmd) {
		switch (cmd) {
			case "Info":
				alert(document.forms.formBuffer.info.value.replace(/\|/gi,"\n"));
				break;
			case "Reload":
				document.forms.formBuffer.submit();
				break;
		}
	}

	function Upload() {
		document.forms.formBuffer.submit();
	}
	
	function PopupRelogin() {	
		opener.Command('Refresh');
		window.close();
	}
	
	function CheckName(str) {
		var re;
		re = /[\\\/:*?"<>|]/gi;
		if (re.test(str)) return false;	
		else return true;
	}	
</script>
<%
	End Sub

	' Writes file listing of the current folder
	Sub List()
		Dim objFolder, virtual, folder
		Dim item, arr
		Dim rowType
		Dim listed

		HtmlHeader appName
		
		on error resume next
		Set objFolder = FSO.GetFolder(targetPath)
		
		If err.Number<>0 Then wexMessage = "Error opening folder !"
		
		virtual = VirtualPath(targetPath)
		folder = right(targetPath, len(targetPath)-len(wexRootPath))
		
%>
<form method="post" action="<%=scriptName%>" name="formBuffer">
	
<input type="hidden" name="command" value>
<input type="hidden" name="parameter" value>
<input type="hidden" name="virtual" value="<%=virtual%>">
<input type="hidden" name="folder" value="<%=folder%>">
<input type="hidden" name="popup" value="false">
	
<table cellspacing="0" cellpadding="2" border="0" width="100%">
	<tr class="titleRow">
		<td align="left" colspan="3">
			&nbsp;<font Face="Arial" Color="White" size="3"><b><%=appName%></b></font>
		</td>
		<td align="right" colspan="2">
			<span class="boldText">v<%=appVersion%> - <%=Date()%></span>&nbsp;
		</td>
	</tr>
	<tr class="lightRow" height="60">
		<td>
			<div style="font-size:12pt; color=black">&nbsp;<img align="absmiddle" border="0" width="45" height="47" src="images/folderopen_big.gif">&nbsp;<span class="blackText"><%=objFolder.Name%></span></div>
			<%If displayPath Then%>
			<!-- <div style="font-size:8pt; color=black">&nbsp;&nbsp;<%=objFolder.path%></div>  -->
			<%End If%>
			<%If virtual<>"" Then%>
				<div style="font-size:8pt;">&nbsp;&nbsp;(<a href="<%=virtual%>" target="_blank" title="Browse the virtual folder"><%=virtual%></a>)</div>
			<%Else%>
				<div style="font-size:8pt;">&nbsp;&nbsp;(<a href="javascript:Command('NoWebAccess');"><%=Session("str")(10)%></a>)</div>
			<%End If%>
		</td>
		<td nowrap>
			<span class="blackText"><%=objFolder.subfolders.count%></span> <%=Session("str")(12)%><br>
			<span class="blackText"><%=objFolder.files.count%></span> <%=Session("str")(13)%>
		</td>
		<td nowrap>
			<%=Session("str")(11)%> <span class="blackText"><%If err.Number<>0 or (not calculateTotalSize) Then Response.Write "N/A" Else Response.Write FormatSize(objFolder.size)%></span>
		</td>
		<td colspan="2" align="center">
			<a href="javascript:Command('Refresh');"><img align="absmiddle" border="0" width="20" height="20" src="images/refresh.gif" alt="<%=Session("str")(0)%>"></a>&nbsp;
			<a href="javascript:Command('NewFile');"><img align="absmiddle" border="0" width="20" height="20" src="images/NewFile.gif" alt="<%=Session("str")(1)%>"></a>&nbsp;
			<a href="javascript:Command('NewFolder');"><img align="absmiddle" border="0" width="20" height="20" src="images/newFolder.gif" alt="<%=Session("str")(2)%>"></a>&nbsp;
			<a href="javascript:Command('Upload');"><img align="absmiddle" border="0" width="20" height="20" src="images/upload.gif" alt="<%=Session("str")(3)%>"></a>&nbsp;
			<a href="javascript:Command('DeleteSelected');"><img align="absmiddle" border="0" width="20" height="20" src="images/del.gif" alt="<%=Session("str")(21)%>"></a>&nbsp;
			<%if showZipUnzip then %>
				<a href="javascript:Command('ZipSelectedFiles');"><img align="absmiddle" border="0" width="20" height="20" src="images/addzip.gif" alt="<%=Session("str")(19)%>"></a>&nbsp;
				<a href="javascript:Command('Unzip');"><img align="absmiddle" border="0" width="20" height="20" src="images/unzip.gif" alt="<%=Session("str")(20)%>"></a>&nbsp;							
			<%end if%>
			<%If wexPassword <> "" Then%>
				<a href="javascript:Command('Logout');"><img align="absmiddle" border="0" width="21" height="20" src="images/logout.png" alt="<%=Session("str")(4)%>"></a>
			<%End If%>
			<br><input name="wexMessage" type="text" class="formClass" size="20" value="<%=server.HTMLEncode(wexMessage)%>" readonly>
		</td>
	</tr>
	<tr class="titleRow">
		<td>&nbsp;<span class="boldText"><%=Session("str")(5)%></span></td>
		<td>&nbsp;<span class="boldText"><%=Session("str")(6)%></span></td>
		<td>&nbsp;<span class="boldText"><%=Session("str")(7)%></span></td>
		<td>&nbsp;<span class="boldText"><%=Session("str")(8)%></span></td>
		<td>&nbsp;<span class="boldText"><%=Session("str")(9)%></span></td>
	</tr>
<%
	rowType = "darkRow"

	If len(targetPath) > len(wexRootPath) Then
%>
	<tr class="<%=rowType%>"><td>&nbsp;<a href="javascript:Command('LevelUp');" title="Up One level"><img align="absmiddle" border="0" width="20" height="20" src="images/folderup.gif"></a>&nbsp;<a href="javascript:Command('LevelUp');" title="Up One level">..</a></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>
<%
		rowType = "lightRow"
	End If
	
	listed = 0
	If (objFolder.subfolders.Count + objFolder.files.Count) = 0 Then
		' Do nothing when error occurs
%>
	<tr class="lightRow">
		<td colspan="5" align="center">No files or folders</td>
	</tr>
<%
	Else
		For each item in objFolder.subfolders
			If showHiddenItems or not item.Attributes and 2 Then
				listed = listed + 1
%>
	<tr class="darkrow">
		<td><input type="checkbox" id="<%=item.name%>_check" name="<%=item.name%>_check">&nbsp;<%=GetIcon(item.Name, true)%>&nbsp;<a href="javascript:Command('OpenFolder',&quot;<%=item.Name%>&quot;);" title="Open Folder"><%=item.Name%></a></td>
		<td>&nbsp;<%If calculateFolderSize Then Response.write FormatSize(item.Size)%></td><td>&nbsp;<%=item.Type%></td>
		<td nowrap>&nbsp;<%=item.DateLastModified%></td>
		<td nowrap>
			&nbsp;			
			<a href="javascript:Command('RenameFolder', &quot;<%=item.Name%>&quot;);"><img align="absmiddle" border="0" width="20" height="20" src="images/ren.gif" alt="<%=Session("str")(25)%>"></a>
			<a href="javascript:Command('DeleteFolder', &quot;<%=item.Name%>&quot;);"><img align="absmiddle" border="0" width="20" height="20" src="images/del.gif" alt="<%=Session("str")(26)%>r"></a>
		</td>
	</tr>
<%
				If rowType = "darkRow" Then rowType = "lightRow" Else rowType = "darkRow"
			End If
		Next

		For each item in objFolder.files
			If showHiddenItems or not item.Attributes and 2 Then
				listed = listed + 1
%>
	<tr class="lightrow">
		<%
			sPath = Virtual
			sPath = Replace(sPath,wexRoot & "/","")
		
		%>
		<td><input type="checkbox" id="<%=item.name%>_check" name="<%=item.name%>_check">&nbsp;<%=GetIcon(item.Name, false)%>&nbsp;<a href="javascript:opener.document.getElementById('<%=Session("FieldName")%>').value='<%=sPath%><%=item.Name%>';self.close();"><%=item.Name%></a></td>
		<td>&nbsp;<%=FormatSize(item.Size)%></td><td>&nbsp;<%=item.Type%></td>
		<td nowrap>&nbsp;<%=item.DateLastModified%></td>
		<td nowrap>
			&nbsp;
			<a href="<%=scriptName & "?precommand=Download&folder=" & Server.URLEncode(folder) & "&file=" & Server.URLEncode(item.Name)%>"><img align="absmiddle" border="0" width="20" height="20" src="images/download.gif" alt="<%=Session("str")(57)%>"></a>
			<a href="javascript:Command('RenameFile', &quot;<%=item.Name%>&quot;);"><img align="absmiddle" border="0" width="20" height="20" src="images/ren.gif" alt="<%=Session("str")(27)%>"></a>
			<a href="javascript:Command('DeleteFile', &quot;<%=item.Name%>&quot;);"><img align="absmiddle" border="0" width="20" height="20" src="images/del.gif" alt="<%=Session("str")(28)%>"></a>			
		</td>
	</tr>
<%
				If rowType = "darkRow" Then rowType = "lightRow" Else rowType = "darkRow"
			End If	
		Next
	End If
	
	'Select ALL Button
%>
	<tr class="titleRow">
		<td colspan="5" class="boldText"><INPUT type="checkbox" onclick="javascript:selectAll(this.form)"><%=Session("str")(22)%></td>
	</tr>
</table>



</form>
<%
		If wexMessage="" Then 
			If (objFolder.subfolders.Count + objFolder.files.Count) <> listed Then
				wexMessage = Session("str")(32) & " " & listed & " of " & (objFolder.subfolders.Count + objFolder.files.Count) & " " & Session("str")(33) & " , " & (objFolder.subfolders.Count + objFolder.files.Count - listed) & " " & Session("str")(34)
			Else
				wexMessage = Session("str")(32) & " " & (objFolder.subfolders.Count + objFolder.files.Count) & " " & Session("str")(33)
			End If
			Response.Write "<script language=""javascript"">document.forms.formBuffer.wexMessage.value='" & wexMessage & "'</script>"
		End If
		
		Set objFolder = Nothing
		
		HtmlCopyright
		HtmlFooter
	End Sub

	' Writes the given error message
	Sub Error(title, message, popup)
		HtmlHeader appName
%>
<table cellpadding="0" cellspacing="0" border="0" align="center" width="300">
	<tr class="titleRow">
		<td>&nbsp;<b><%=Session("str")(35)%></b></td>
	</tr>
	<tr class="lightRow">
		<td>
			<table cellpadding="0" cellspacing="5" border="0">
				<tr>
					<td valign="top"><img width="32" height="32" border="0" align="absmiddle" src="images/error.png"></td>
					<td><b><%=title%>:</b><br><%=message%></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr class="titleRow" align="center">
		<td>
			<%If popup Then%>
			<a href="javascript:this.close();"><%=Session("str")(18)%></a>
			<%Else%>
			<a href="javascript:history.back();"><%=Session("str")(36)%></a>
			<%End If%>
		</td>
	</tr>
</table>
<%
		HtmlFooter
		DestroyApp()
	End Sub
	
	' WebExplorer Lite login screen
	Sub	Login()
		If Request.Form("command") = "Login" Then
			If Request.Form("pwd") = wexPassword Then
				Session(wexId & "Login") = true
				Exit Sub
			Else
				wexMessage = "Wrong password!"
			End If
		End If
		
		HtmlHeader appName
		If(wexMessage<>"") Then Response.Write "<script language=""javascript"">alert('" & wexMessage & "');</script>"
%>
<form name="formBuffer" method="post" action="<%=scriptName%>" onSubmit="javascript:return(Check());">
<table border="0" cellspacing="0" cellpadding="0" width="400" align="center">
	<tr><td><br><br><br></td></tr>
	<tr><td>
		<table border="0" cellspacing="0" cellpadding="0" width="100%">
			<tr class="titleRow">
				<td align="left">
					&nbsp;<span class="boldText">Login</span>
				</td>
			</tr>
		</table>
		<table border="0" cellspacing="0" cellpadding="0" width="100%">
			<tr align="center" class="lightRow">
				<td>
					<br>
					<span class="boldText">Welcome to <%=appName%> v<%=appVersion%></span>
					<br><br>
					<table cellspacing="0" cellpadding="5" border="0" class="loginRow">
						<tr>
							<td align="left">&nbsp;<span class="boldText">Password</span></td>
						</tr>
						<tr>
							<td align="center"><input type="password" class="formClass" name="pwd" value size="21"></td>
						</tr>
						<tr>
							<td align="right"><input type="submit" name="submitter" value="Login" class="formClass"></td>
						</tr>
					</table>
					<br><br><br>
				</td>
			</tr>
		</table>
		<table border="0" cellspacing="0" cellpadding="0" width="100%">
			<tr class="titleRow">
				<td align="center">&nbsp;</td>
			</tr>
		</table>
	</td></tr>
</table>
<input type="hidden" name="command" value="Login">
</form>
<script language="javascript">document.forms.formBuffer.pwd.focus();</script>
<%	
		HtmlFooter
		DestroyApp() 
	End Sub
	
	' Relogin message for the popup windows
	Sub PopupRelogin()
		HtmlHeader appName
%>
		<div style="COLOR: white; FONT-FAMILY: Verdana,Tahoma,Arial,Helvetica; FONT-SIZE: 10pt; FONT-WEIGHT: bold;">
		<%=appName%> session is destroyed, please <a href="javascript:PopupRelogin();">relogin</a>.
		</div>
<%
		HtmlFooter
		DestroyApp() 
	End Sub

	' Checks if there is a valid login
	Function Secure()
		If wexPassword = "" Then
			Secure = true
		Else
			If Session(wexId & "Login") Then Secure = true Else Secure = false
		End If
	End Function
		
	' Logs out from WebExplorer Lite
	Sub Logout()
		Session.Abandon()
		Login
	End Sub
		
	' Returns the icon of the file
	Function GetIcon(fileName, isFolder)
		Dim ext

		If isFolder Then
			GetIcon = "<a href=""javascript:Command('FolderDetails', &#34;" & fileName & "&#34;);""><img align=absmiddle border=0 width=16 height=16 src=""./images/folder.gif"" alt=""" &  Session("str")(14) & """></a>"
		Else
			ext = FSO.GetExtensionName(fileName)
			
			re.IgnoreCase = true
			re.Pattern = "^" & ext & ",|," & ext & ",|," & ext & "$"

			If re.test(editableExtensions) Then	
				
				GetIcon = "<a href=""javascript:Command('Edit', &#34;" & fileName & "&#34;);""><img align=absmiddle border=0 width=16 height=16 src=""./images/t_" & Cstr(Ext) &  ".gif"" alt=""Text file - Click to edit and learn details""></a>"

			ElseIf re.test(viewableExtensions) Then

				GetIcon = "<a href=""javascript:Command('View', &#34;" & fileName & "&#34;);""><img align=absmiddle border=0 width=16 height=16 src=""./images/t_" & Cstr(Ext) &  ".gif"" alt=""Picture file - Click to view and learn details""></a>"

			Elseif re.test(otherExtensions) Then
				GetIcon = "<a href=""javascript:Command('FileDetails', &#34;" & fileName & "&#34;);""><img align=absmiddle border=0 width=16 height=16 src=""./images/t_" & Cstr(Ext) &  ".gif"" alt=""File - Click to learn details""></a>"			
			else
				GetIcon = "<a href=""javascript:Command('FileDetails', &#34;" & fileName & "&#34;);""><img align=absmiddle border=0 width=16 height=16 src=""./images/t_blank.gif"" alt=""File - Click to learn details""></a>"
			End If
		End If
	End Function
	
	' Formats given size in bytes,KB,MB and GB
	Function FormatSize(givenSize)
		If (givenSize < 1024) Then
			FormatSize = givenSize & " B"
		ElseIf (givenSize < 1024*1024) Then
			FormatSize = FormatNumber(givenSize/1024,2) & " KB"
		ElseIf (givenSize < 1024*1024*1024) Then
			FormatSize = FormatNumber(givenSize/(1024*1024),2) & " MB"
		Else
			FormatSize = FormatNumber(givenSize/(1024*1024*1024),2) & " GB"
		End If
	End Function

	' Adds given type of the slash to the end of the path if required
	Function FixPath(path, slash)
		If Right(path, 1) <> slash Then
            FixPath = path & slash
        Else
			FixPath = path
        End If
	End Function

	' Converts the given path to physical path
	Function RealizePath(path)
		Dim fpath
		fpath = replace(path,"/","\")
		If left(fpath,1) = "\" Then 'Virtual path
			on error resume next
			RealizePath = server.MapPath(fpath)
			If err.Number<>0 Then RealizePath = fpath 'Possibly network path
		Else 'Physical Path
			RealizePath = fpath
		End If
		RealizePath = FixPath(RealizePath, "\")
	End Function	

	' Converts the given path to virtual path
	Function VirtualPath(path)
		Dim webRoot, fpath
		webRoot = FixPath(server.MapPath("/"),"\")
		fpath = FixPath(path,"\")
		VirtualPath = ""
		If left(wexRoot,1) = "/" Then
			VirtualPath = FixPath(wexRoot, "/")
			VirtualPath = VirtualPath & right(fpath, len(fpath) - len(wexRootPath))
			VirtualPath = replace(VirtualPath, "\", "/")
			VirtualPath = FixPath(VirtualPath,"/")
		ElseIf left(lcase(fpath), len(webRoot)) = lcase(webRoot) Then
			VirtualPath = "/" & right(fpath, len(fpath) - len(webRoot))
			VirtualPath = replace(VirtualPath, "\", "/")
			VirtualPath = FixPath(VirtualPath,"/")
		End If
	End Function

	'Maps the given path according to the root path
	Function WexMapPath(path)
		If SecurePath(path) Then WexMapPath = FixPath(wexRootPath & path, "\") Else Error "Security Error", "Relative path syntax is forbidden for security reasons.", false
	End Function
		
	' Checks against relative path syntax (. or .. injection)
	Function SecurePath(path)
		Dim fpath
		fpath = replace(path,"/","\")
		
		If fpath="." Then fpath=".\"

		re.IgnoreCase = false
		re.Pattern = "^\.\.$|^\.\.\\|\\\.\.\\|\\\.\.$"
		re.Pattern = re.Pattern & "|^\.\\|\\\.\\|\\\.$"
		
		If re.Test(fpath) Then SecurePath=false Else SecurePath=true
	End Function
	
	' Makes sure that given file name does not contain path info
	Function SecureFileName(name)
		SecureFileName = replace(name,"/","?")
		SecureFileName = replace(SecureFileName,"\","?")
	End Function

	' Checks if the extension of the given file name is allowed
	Function CheckExtension(fileName)
		Dim allow
		Dim re, match, extension
	
		If monitoredExtensions<>"" Then
			Set re = new regexp

			re.IgnoreCase = true
			re.Global = false
			re.Pattern = "\.(\w+)$"
			Set match = re.Execute(fileName)
			If match.Count<>0 Then
				extension = match(0).SubMatches(0)
				re.Pattern = "^" & extension & ",|," & extension & ",|," & extension & "$" & "|^" & extension & "$"
				If re.test(monitoredExtensions) Then allow = false Else allow = true
			Else
				allow = true
			End If
			
			Set re = Nothing

			If denyMonitored Then
				CheckExtension = allow
			Else
				CheckExtension = (not allow)
			End If
		Else
			CheckExtension = true
		End If
	End Function

	' Creates a folder or a file
	Function CreateItem()
		Dim itemType, itemName, itemPath
		itemType = Request.Form("command")
		itemName = SecureFileName(Request.Form("parameter"))
		itemPath = targetPath & itemName

		on error resume next
		
		Select Case itemType
			Case "NewFolder"
				If FSO.FolderExists(itemPath) = false and FSO.FileExists(itemPath) = false Then 
					FSO.CreateFolder(itemPath)
					If err.Number <> 0 Then 
						wexMessage = "Unable to create the folder """ & itemName & """, an error occured..." 
					Else
						wexMessage = "Created the folder """ & itemName & """..."
					End If
				Else
					wexMessage = "Unable to create the folder """ & itemName & """, there exists a file or a folder with the same name..."
				End If
			Case "NewFile"
				If FSO.FolderExists(itemPath) = false and FSO.FileExists(itemPath) = false Then 
					If CheckExtension(itemName) Then 
						FSO.CreateTextFile(itemPath)
					Else
						err.Raise 1
					End If
					
					If err.Number <> 0 Then 
						wexMessage = "Unable to create the file """ & itemName & """, an error occured..."
					Else
						wexMessage = "Created the file """ & itemName & """..."
					End If
				Else 
					wexMessage = "Unable to create the file """ & itemName & """, there exists a file or a folder with the same name..."
				End IF
		End Select
	End Function
	
	' Deletes a folder or a file
	Function DeleteItem()
		Dim itemType, itemName, itemPath
		
		Dim objFolder, virtual, folder
		Dim item, arr

		
		on error resume next
		Set objFolder = FSO.GetFolder(targetPath)
			
		
		itemType = Request.Form("command")
		itemName = SecureFileName(Request.Form("parameter"))
		itemPath = targetPath & itemName

		on error resume next
		
		Select Case itemType
			Case "DeleteFolder"
				FSO.DeleteFolder itemPath, true
				If err.Number <> 0 Then 
					wexMessage = Session("str")(30) & " " & itemName  
				Else
					wexMessage = Session("str")(31) & " " & itemName  
				End If
			Case "DeleteFile"
				FSO.DeleteFile itemPath, true
				If err.Number <> 0 Then 
					wexMessage = Session("str")(30) & " " & itemName  				
				Else
					wexMessage = Session("str")(31) & " " & itemName  
				End If
			Case "DeleteSelected"
				
				For each item in objFolder.files					
					if CBool(Request.Form(item.name+"_check")) then					
						itemPath = targetPath & item.name
						FSO.DeleteFile itemPath, true
						If err.Number <> 0 Then 
							wexMessage = Session("str")(30) & " " & itemName  
						Else
							exMessage = Session("str")(31) & " " & itemName
						End If						
						
					end if
				next
				
				For each item in objFolder.subfolders					
					if CBool(Request.Form(item.name+"_check")) then		
						itemPath = targetPath & item.name
						FSO.DeleteFolder itemPath, true
						If err.Number <> 0 Then 
							wexMessage = Session("str")(30) & " " & itemName  
						Else
							wexMessage = Session("str")(31) & " " & itemName  
						end if
					End If	
				next						
			
		End Select
		
		Set objFolder=nothing
	End Function

	' Renames a folder or a file
	Function RenameItem()
		Dim item, itemType, itemName, itemPath
		Dim param, newName
		itemType = Request.Form("command")
		param = split(Request.Form("parameter"), "|")
		itemName = SecureFileName(param(0))
		newName = SecureFileName(param(1))
		itemPath = targetPath & newName

		on error resume next
		
		Select Case itemType
			Case "RenameFolder"
				If FSO.FolderExists(itemPath) = false and FSO.FileExists(itemPath) = false Then 
					itemPath = targetPath & itemName
					Set item = FSO.GetFolder(itemPath)
					item.Name = newName
					If err.Number <> 0 Then 
						wexMessage = "Unable to rename the folder """ & itemName & """, an error occured..." 
					Else
						wexMessage = "Renamed the folder """ & itemName & """ to """ & newName & """..."
					End If
				Else
					wexMessage = "Unable to rename the folder """ & itemName & """, there exists a file or a folder with the new name """ & newName & """..."
				End If
			Case "RenameFile"
				If FSO.FolderExists(itemPath) = false and FSO.FileExists(itemPath) = false Then 
					If CheckExtension(newName) Then 
						itemPath = targetPath & itemName
						Set item = FSO.GetFile(itemPath)
						item.Name = newName
					Else
						err.Raise 1
					End If

					If err.Number <> 0 Then 
						wexMessage = "Unable to rename the file """ & itemName & """, an error occured..." 
					Else
						wexMessage = "Renamed the file """ & itemName & """ to """ & newName & """..."
					End If
				Else
					wexMessage = "Unable to rename the file """ & itemName & """, there exists a file or a folder with the new name """ & newName & """..."
				End If
		End Select
		
		Set item = Nothing
	End Function
		
	' WebExplorer Lite Editor
	Sub Editor()
		Dim fileName, filePath, file
		
		on error resume next

		Select Case Request.Form("subcommand")
			Case "Save", "SaveAs"
				fileName = SecureFileName(Request.Form("parameter"))
				filePath = targetPath & fileName
				
				If CheckExtension(fileName) Then 
					Set file = FSO.OpenTextFile (filePath, 2, true, encoding)
					If (err.Number<>0) Then 
						wexMessage = "Can not write to the file """ & fileName & """, permission denied!"
						err.Clear
					Else
						file.write Request.Form("content")
					End If
					Set file = Nothing
				Else
					wexMessage = "Can not write to the file """ & fileName & """, extension not allowed!"
				End If

				Set file = FSO.OpenTextFile (filePath, 1, false, encoding)
			Case Else
				fileName = SecureFileName(Request.Form("parameter"))
				filePath = targetPath & fileName
				
				If not FSO.FileExists(filePath) Then
					wexMessage = "The file """ & fileName & """ does not exist"
					Set file = FSO.CreateTextFile (filePath, false)
					If err.Number<>0 Then 
						wexMessage = wexMessage & ", also unable to create new file."
						err.Clear 
					Else
						wexMessage = wexMessage & ", created new file."
					End If
				Else
					Set file = FSO.OpenTextFile (filePath, 1, false, encoding)
					If err.Number<>0 Then 
						wexMessage = "Can not read from the file """ & fileName & """, permission denied!"
						err.Clear 
					End If
				End If
		End Select

		HtmlHeader appName
		If(wexMessage<>"") Then Response.Write "<script language=""javascript"">alert('" & wexMessage & "');</script>"
%>
<form name="formBuffer" method="post" action="<%=scriptName%>">
		<table border="0" cellspacing="0" cellpadding="0" width="100%">
			<tr class="DarkRow">
				<td align="left">
					&nbsp;<span class="blacktext"><%=Session("str")(37)%> - <%=fileName%></span>
				</td>
				<td align="right">
					<b><%=Session("str")(38)%>:</b>
					<select name="encoding" class="formClass" onChange="EditorCommand('Encoding')">
						<option value="-2" <%If encoding=-2 Then Response.Write " selected"%>>Default</option>
						<option value="-1" <%If encoding=-1 Then Response.Write " selected"%>>Unicode</option>
					</select>
				</td>
			</tr>
		</table>
		<table border="0" cellspacing="0" cellpadding="0" width="100%" height="90%">
			<tr align="center" class="lightRow">
				<td valign="middle">
<textarea name="content" class="formClass" rows="22" cols="46" style="width:580; height:370;" wrap="off">
<%=Server.HTMLEncode(file.ReadAll)%></textarea>
				</td>
			</tr>
		</table>
		<table border="0" cellspacing="0" cellpadding="0" width="100%">
			<tr class="DarkRow">
				<td align="center">					
					<a href="javascript:EditorCommand('Save');"><%=Session("str")(39)%></a> | <a href="javascript:EditorCommand('SaveAs');"><%=Session("str")(40)%></a> | <a href="javascript:EditorCommand('Reload');"><%=Session("str")(0)%></a> | <a href="javascript:EditorCommand('Info');"><%=Session("str")(17)%></a> | <a href="javascript:this.close();"><%=Session("str")(18)%></a>
				</td>
			</tr>
		</table>
<%
		Set file = Nothing
		Set file = FSO.GetFile (filePath)
%>
<input type="hidden" name="command" value="Edit">
<input type="hidden" name="subcommand" value>
<input type="hidden" name="parameter" value="<%=fileName%>">
<input type="hidden" name="folder" value="<%=Request.Form("folder")%>">
<input type="hidden" name="info" value="Size: <%=FormatSize(file.Size)%>|Type: <%=file.Type%>|Created: <%=file.DateCreated%>|Last Accessed: <%=file.DateLastAccessed%>|Last Modified: <%=file.DateLastModified%>">
<input type="hidden" name="popup" value="true">
</form>
<%
		Set file = Nothing

		HtmlFooter
		DestroyApp() 
	End Sub

	' WebExplorer Lite Viewer
	Sub Viewer()
		Dim filePath, file

		filePath = targetPath & Request.Form("parameter")
		If not FSO.FileExists(filePath) Then Error "Viewer Error", "File not found. Please refresh the listing to see if the file actually exists.", true
		
		on error resume next
		Set file = FSO.GetFile(filePath)

		HtmlHeader appName
%>
<form name="formBuffer" method="post" action="<%=scriptName%>">
		<table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%">
			<tr class="DarkRow" height="20">
				<td align="left">
					&nbsp;<span class="boldText"><%=Session("str")(15)%></span> - <%=file.Name%>
				</td>
			</tr>
			<tr align="center" class="lightRow">
				<td valign="middle">
					<img src="<%=Replace(wexRoot,"\","/") & "/" & Replace(Request.Form("folder"),"\","/") & file.Name%>">
				</td>
			</tr>
			<tr class="DarkRow" height="20">
				<td align="center">					
					<a href="javascript:ViewerCommand('Reload');"><%=Session("str")(16)%></a> | <a href="javascript:ViewerCommand('Info');"><%=Session("str")(17)%></a> | <a href="javascript:this.close();"><%=Session("str")(18)%></a>
				</td>
			</tr>
		</table>
<input type="hidden" name="command" value="View">
<input type="hidden" name="subcommand" value="Refresh">
<input type="hidden" name="parameter" value="<%=file.Name%>">
<input type="hidden" name="folder" value="<%=Request.Form("folder")%>">
<input type="hidden" name="info" value="Size: <%=FormatSize(file.Size)%>|Type: <%=file.Type%>|Created: <%=file.DateCreated%>|Last Accessed: <%=file.DateLastAccessed%>|Last Modified: <%=file.DateLastModified%>">
<input type="hidden" name="popup" value="true">
</form>
<%
		Set file = Nothing
		HtmlFooter
		DestroyApp() 
	End Sub

	' File/Folder Details
	Sub Details()
		Dim fileName, filePath, file
		
		on error resume next
		fileName = Request.Form("parameter")
		filePath = targetPath & fileName

		HtmlHeader appName
%>
<form name="formBuffer" method="post" action="<%=scriptName%>">
		<table border="0" cellspacing="0" cellpadding="0" width="100%">
			<tr class="titleRow">
				<td align="left">
					&nbsp;<span class="boldText">Details</span> - <%=fileName%>
				</td>
			</tr>
		</table>
		<table border="0" cellspacing="0" cellpadding="0" width="100%" height="80%">
			<tr align="center" class="lightRow">
				<td valign="middle">
<%
		If Request.Form("command") = "FileDetails" Then
				Set file = FSO.GetFile (filePath)
		Else
				Set file = FSO.GetFolder (filePath)
		End If
%>
				<table border="0" cellspacing="5" cellpadding="0">
					<tr><td><span class="boldText">Size:</span></td><td><%=FormatSize(file.Size)%></td></tr>
					<tr><td><span class="boldText">Type:</span></td><td><%=file.Type%></td></tr>
					<tr><td><span class="boldText">Created:</span></td><td><%=file.DateCreated%></td></tr>
					<tr><td><span class="boldText">Last Accessed:</span></td><td><%=file.DateLastAccessed%></td></tr>
					<tr><td><span class="boldText">Last Modified:</span></td><td><%=file.DateLastModified%></td></tr>
				</table>
<%
		Set file = Nothing
%>
				</td>
			</tr>
		</table>
		<table border="0" cellspacing="0" cellpadding="0" width="100%">
			<tr class="titleRow">
				<td align="center">					
					<a href="javascript:this.close();">Close</a>
				</td>
			</tr>
		</table>
<input type="hidden" name="command" value="<%=Request.Form("command")%>">
<input type="hidden" name="parameter" value="<%=fileName%>">
<input type="hidden" name="folder" value="<%=Request.Form("folder")%>">
<input type="hidden" name="popup" value="true">
</form>
<%
		HtmlFooter
		DestroyApp() 
	End Sub
	
	' Uploads a file
	Sub Upload(process)
		Dim fileTransfer, result

		on error resume next
		Set fileTransfer = New pluginFileTransfer
		If err.number<>0 Then Error "File Transfer Plugin Error", "Plugin cannot be initialized. Please make sure that the components required by the plugin is installed on the server.", true

		If process Then targetPath = WexMapPath(Request.QueryString("folder"))
		
		HtmlHeader appName
%>
		<table border="0" cellspacing="0" cellpadding="0" width="100%">
			<tr class="DarkRow">
				<td align="left">
					&nbsp;<span class="blacktext"><%=Session("str")(44)%> - <%=FSO.GetBaseName(targetPath)%></span>
				</td>
			</tr>
		</table>
		<table border="0" cellspacing="0" cellpadding="0" width="100%" height="80%">
			<tr align="center" class="lightRow">
				<td valign="middle">
					<span class="blacktext">
<%
		' Actual upload process
		If process Then
			fileTransfer.path = targetPath
			
			result = fileTransfer.Upload()
			
			Select Case result
				Case 0
					Response.Write fileTransfer.uploadedFileName & " " & Session("str")(50) & "<br><br>"
					Response.Write FormatSize(fileTransfer.uploadedFileSize) & " (" & fileTransfer.uploadedFileSize & " bytes) written<br>"
					Response.Write "Content type: " & fileTransfer.contentType
					Response.Write "<script language=""javascript"">opener.Command('Refresh');</script>"
				Case 1
					Response.Write Session("str")(46)
				Case 2
					Response.Write Session("str")(47)
				Case 3
					Response.Write fileTransfer.uploadedFileName & Session("str")(46)
				Case 4
					Response.Write Session("str")(49)
			End Select
%>
					</span>
					<form name="formBuffer" method="post" action="<%=scriptName%>">
						<input type="hidden" name="command" value="Upload">
						<input type="hidden" name="folder" value="<%=Request.QueryString("folder")%>">
					</form>
<%	
		Else
%>
					<form enctype="multipart/form-data" name="formBuffer" method="post" action="<%=scriptName%>?precommand=ProcessUpload&amp;folder=<%=server.URLEncode(Request.Form("folder"))%>&amp;popup=true">
						<input type="file" name="file" class="formClass">
					</form>
<%
		End If
%>					
				</td>
			</tr>
		</table>
		<table border="0" cellspacing="0" cellpadding="0" width="100%">
			<tr class="DarkRow">
				<td align="center">					
					<a href="javascript:Upload();"><%=Session("str")(45)%></a> | <a href="javascript:this.close();"><%=Session("str")(18)%></a>
				</td>
			</tr>
		</table>
<%
		Set fileTransfer = Nothing

		HtmlFooter
		DestroyApp() 
	End Sub
	
	' Downloads a file
	Sub Download()
		Dim fileTransfer, result
		
		on error resume next
		Set fileTransfer = New pluginFileTransfer
		If err.number<>0 Then Error "File Transfer Plugin Error", "Plugin cannot be initialized. Please make sure that the components required by the plugin is installed on the server.", false
		
		fileTransfer.path = WexMapPath(Request.QueryString("folder"))
		
		result = fileTransfer.Download(Request.QueryString("file"))
		
		Select Case result
			Case 0
				'Success
			Case 1
				Error "Download Error", "File not found. Please refresh the listing to see if the file actually exists.", false
			Case 2
				Error "Download Error", "File cannot be read. Please make sure that you have read permission on the file.", false
		End Select
		
		Set fileTransfer = Nothing

		DestroyApp() 
	End Sub
	
	
'########################
SUB MoveItems(f,Target)
'########################
	If Right(Target,1)<>"\" Then Target=Target & "\"
	'### Check if target RF is not full, but allow Bin
	If Session("UseRootfolders") Then 
		If Instr(1,Target,Session("FMRecyclerName"),1)<>1 Then 
			TargetRFNum=GetRFNum(Target)
			CheckRootfolder(TargetRFNum)
			If Session("IsReadOnly") Then ShowError "The target folder is full or read-only!"
		End If
	End If
	tArr=Split(f,", ")
	For i=0 to Ubound(tArr)
		If Application("Debugging")=False Then On Error resume next
		tArr(i)=decPath(tArr(i))
		pf=fso.GetParentFoldername(Lcase(tArr(0)))& "\" 
		If pf = Lcase(Target) Then IsSameFolder=True Else IsSameFolder=False
		If Right(tArr(i),1)="\" Then
			If Instr(1,target,tArr(i),1)=1 Then ShowError("You can't move this folder: destination is a subfolder of the source!")
			If Session("Settings")(14) AND NOT IsSameFolder Then
				fn=fso.GetBaseName(tArr(i)) & "." & fso.getExtensionName(tArr(i))
				If fso.FolderExists(Target & fn) Then
					fso.copyfolder Left(tArr(i),len(tArr(i))-1),Target & fn, Session("Settings")(23)
					If err=0 then fso.deleteFolder Left(tArr(i),Len(tArr(i))-1), Session("Settings")(23)
				Else
					fso.movefolder Left(tArr(i),len(tArr(i))-1) ,Target
				End If
			End If
		Else
		 	If NOT IsSameFolder Then
				fn=fso.getFilename(tArr(i))
				If fso.FileExists(Target & fn) Then fso.deletefile Target & fn, Session("Settings")(23)
				fso.movefile tArr(i),Target
			End If
		End If
		If Application("LogLevel")>1 Then WriteLogLine("Move " & tArr(i) & " to " & Target)
	Next
	If err<>0 Then Call ShowError(Session("Str")(38) & " -> " & RelativePath(target))
	If Instr(1,Target,Session("FMRecyclerName"),1)>0 Then CountRecyclerItems
	Session("NumInQueue")=0
	CheckRootfolder(Session("CurRFNum"))
	Response.redirect "fileman.asp"
End SUB


'########################
SUB CopyItems(f,Target)
'########################
	If Right(target,1)<>"\" Then target=target & "\"
	'### Check if target RF is not full
	If Session("UseRootfolders") Then 
		TargetRFNum=GetRFNum(Target)
		CheckRootfolder(TargetRFNum)
		If Session("IsReadOnly") Then ShowError "The target folder is full or read-only!"
	End If
	tArr=Split(f,", ")
	For i=0 to Ubound(tArr)
		If Application("Debugging")=False Then On Error resume next
		tArr(i)=decPath(tArr(i))
		pf=fso.GetParentFoldername(Lcase(tArr(0)))& "\" 
		If pf = Lcase(Target) Then IsSameFolder=True Else IsSameFolder=False
		If Right(tArr(i),1)="\" Then
			If Instr(1,Target,tArr(i),1)=1 AND NOT IsSameFolder Then ShowError("You can't copy this folder: destination is a subfolder of the source!")
			If Session("Settings")(14) Then
				fn=fso.GetBaseName(tArr(i)) & "." & fso.getExtensionName(tArr(i))
			 	If IsSameFolder Then
					If Instr(fn,"Copy of ")=1 OR Instr(fn,"Copy (")=1 Then fn=Mid(fn,Instr(fn, " of ")+4)
					If Not fso.FolderExists(Target & fn) Then
						fn=fn		
					ElseIf Not fso.FolderExists(Target & "Copy of " & fn) Then
						fn="Copy of " & fn
					Else
						tn="Copy (1) of " & fn
						n=0
						While fso.FolderExists(Target & tn) and n<99
							n=n+1
							tn="Copy (" & n & ") of " & fn
						Wend
						fn=tn
					End If
				End If
				fso.copyfolder Left(tArr(i),len(tArr(i))-1),target & fn, Session("Settings")(23)
			End If
		Else
			fn=fso.getfilename(tArr(i))
		 	If IsSameFolder Then
				If Instr(fn,"Copy of ")=1 OR Instr(fn,"Copy (")=1 Then fn=Mid(fn,Instr(fn, " of ")+4)
				If Not fso.FileExists(Target & fn) Then
					fn=fn
				ElseIf Not fso.FileExists(Target & "Copy of " & fn) Then
					fn="Copy of " & fn
				Else
					tn="Copy (1) of " & fn
					n=0
					While fso.FileExists(Target & tn) and n<99
						n=n+1
						tn="Copy (" & n & ") of " & fn
					Wend
					fn=tn
				End If
			End If
			fso.copyfile tArr(i),Target & fn,Session("Settings")(23)
		End If
		If Application("LogLevel")>1 Then WriteLogLine("Copy " & tArr(i) & " to " & Target & fn)
	Next
	If err<>0 Then Call ShowError(Session("Str")(148) & " " & Session("Str")(144) & " " & RelativePath(target))
	If Instr(1,Target & f,Session("FMRecyclerName"),1)>0 Then CountRecyclerItems
	CheckRootfolder(Session("CurRFNum"))
	Session("NumInQueue")=0
	Response.redirect "fileman.asp"
End SUB	
	


	
'########################
SUB ZipSelectedFiles()
'########################

	Dim itemType, itemName, itemPath
		
	Dim objFolder, virtual, folder
	Dim item, arr

		
	on error resume next
	Set objFolder = FSO.GetFolder(targetPath)
			
		
	itemType = Request.Form("command")
	Target = SecureFileName(Request.Form("parameter"))
	Target = targetPath & Target
	itemPath = targetPath & itemName
	
	Set oZip = Server.CreateObject ("FathZIP.FathZIPCtrl.1")
	If UCase(fso.GetExtensionName(Target))<> "ZIP" Then Target=Target & ".ZIP"
	If NOT fso.fileexists(Target) Then oZip.CreateZip Target, "" Else oZip.OpenZip Target 	
	oZip.Basepath=targetPath	
	oZip.PreservePaths=true	
	
	For each item in objFolder.files					
		if CBool(Request.Form(item.name+"_check")) then					
				itemPath = targetPath & item.name
				oZip.AddFile itemPath, ""
		end if
	next			
	
	set item = nothing
			
	For each item in objFolder.subfolders						
		if CBool(Request.Form(item.name+"_check")) then					
			itemPath = targetPath & item.name
			oZip.AddFile itemPath, ""
			
		end if
	next			
	
	oZip.Close

End SUB

'########################
SUB UnZip()
'########################
	Dim itemType, itemName, itemPath
		
	Dim objFolder, virtual, folder
	Dim item, arr

		
	on error resume next
	Set objFolder = FSO.GetFolder(targetPath)
			
		
	itemType = Request.Form("command")
	itemName = SecureFileName(Request.Form("parameter"))
	itemPath = targetPath & itemName

	on error resume next	
	
	Set oZip = Server.CreateObject ("FathZIP.FathZIPCtrl.1")
	
	For each item in objFolder.files					
		if CBool(Request.Form(item.name+"_check")) then					
			itemPath = targetPath & item.name
			oZip.Basepath=targetPath
			oZip.PreservePaths=True
			
			
			
			If UCase(fso.GetExtensionName(item.name))="ZIP" Then			
				oZip.OpenZip itemPath
				For n= 0 to oZip.FileCount-1
						If oZip.Filename(n) Then 
							oZip.ExtractFile oZip.Filename(n)
						end if
					
				Next
				oZip.Close			
			End If
							
						
		end if
	next
	Set oZip = Nothing	
	
End SUB


	
' ------------------------------------------------------------
%>

