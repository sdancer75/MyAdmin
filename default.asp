<%@ Language=VBScript%>



<!-- #include file="include/functions.asp" -->
<%
intLocale = SetLocale(1032)
'intLocale = SetLocale(el)

if Request.Form("action")="logon" then
	if CheckUsername( CStr(request.form("postusername")),CStr(request.form("postpassword")),CStr(request.form("remember")) ) then
		Response.Redirect sLoginStartPage	
	end if
end if

	Response.CacheControl = "no-cache"
	Response.Expires = -1		
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head><title>MyAdmin copyright (c) Paradox Interactive</title>
<link href="include/style.css" rel="stylesheet" type="text/css">
<!-- #include file="include/METATags.asp" -->
</head>
<div align="center">
	
	<%=sAuthAccessTitle%>
</div> 


<div align="center">
<table cellPadding="0" cellSpacing="0" width="300" align="center" bgcolor="#cccccc" bordercolor="White">  
<tr> 
       <td align=middle height=50 valign=middle>
		<b><font size=+1>Login</font></b></td>
		
</tr>
<tr>	
	<td border="2" width="300" valign="middle" bordercolor="black" align="center">
		
		<form ENCTYPE="application/x-www-form-urlencoded" method="POST" id="form1" name="form1" action="default.asp">
		   <input type="hidden" name="action" value="logon">
		   <span class="smalltext" align="center"><%=html_username%></span>
		   <input type="text" id="text1" name="postusername" maxlength="20" size="30"  value="<%=LoadCookie("qo_username")%>"><br>
		   <span class="smalltext" align="left"><%=html_password%></span>
		   <input type="password" id="password1" name="postpassword" maxlength="20" size="30"  value="<%=LoadCookie("qo_password")%>"><br>
		   <INPUT type="checkbox" id=remember name=remember> Remember password
		   <br><br>
		   <div align="center">
		   <input type="submit" value="<%=html_submit%>" id="button1" name="button1" >		   
		   </div>
	</form>		   
	</td>
</tr>
<tr>
	<td align=center height=20>
		<font color=red><%=sResult%></font>
	</td>
</tr>
<tr>
	<td height=20>			
		   <div align="left"><a href="javascript: history.back(-1);"><img SRC="images/prev.gif" ALIGN="Absmiddle" BORDER="0" hspace="3" VSPACE="0" WIDTH="14" HEIGHT="14">Go Back</a></div>
	
	</td>
</tr>
</table>
</div>

<!-- #include file="include/footer.asp" -->



</body>
</html>