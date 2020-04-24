<!-- #include file="adovbs.inc" -->
<!-- #include file="WebSiteAccount.asp" -->

<%



	Const tbl_UsersTable					="Users"
	Const f_Username						="username"
	Const f_Password						="password"
	Const f_RelationField					="CSN"					


	Const html_username					="Username"
	Const html_password					="Password"
	Const html_submit						="submit"

	Const msg_NoValidLogin				="No Valid Username or Password"
	Const sAuthAccessTitle				=""
	Const sLoginStartPage					="admin.asp"

	Const sMainMenuTitle					="Administrator's page"


	Dim MyConn,rs,rsMaster
	Dim dbconnect,sResult
	Dim LoginAdmin


	Set MyConn = Server.CreateObject("ADODB.Connection")
	dbconnect = "Driver={SQL Server}; Server=" & sMSSQLServerAddress & "; Database=" &  sMSSQLInitialCatalog & "; Uid=" & sMSSQLUsername & "; Pwd=" & sMSSQLPassword
	MyConn.Open  dbconnect
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rsMaster=Server.CreateObject("ADODB.Recordset")
	
	
	Response.CacheControl = "no-cache"
	Response.Expires = -1		





Function ValidateField(byval sSQLstring)
   ' DESCRIPTION: Escapes out HTML-unsave characters
   if isNull(sSQLString) then sSqlString = ""
   sSQLstring= trim(sSQLstring)
   sSQLstring= Replace (sSQLstring, CHR(34), "&quot;")
   sSQLstring= Replace (sSQLstring, "<", "&lt;")
   sSQLstring= Replace (sSQLstring, ">", "&gt;")
   'ValidateField = server.HTMLEncode(sSQLString)
ValidateField = sSQLString
End Function

Function LoadCookie(byval cookiename)

	LoadCookie=Request.Cookies(cookiename)

End Function


Function ValidateSQL(byval sSQLstring)
   ' DESCRIPTION: Properly formats a string for use in an SQL statement. Preserves value
   if isNull(sSQLString) then sSqlString = ""
   sSQLstring= trim(sSQLstring)
   sSQLstring= Replace (sSQLstring, "'", "''")
   ValidateSQL = sSQLstring
End Function


Function CheckUsername (byval sUsername, byval sPassword, byval sRemember)
  
  dim SQL  
 
  sResult=""
  if validateSQL(sUsername)= "" then
	CheckUserName=false
	sResult = CStr(msg_NoValidLogin)	
	exit function
  end if
  	   
  SQL = "select * from [" & CStr(tbl_UsersTable) & "] where [" & CStr(f_UserName) & "]='" & validateSQL(sUsername) & "';"
 
  rs.open SQL,MyConn, adOpenStatic,adLockOptimistic,adCmdText
  'rs.Open "Select * from users",
  
	
  if rs.EOF=false then   
  
		if rs(CStr(f_Password)) = trim(sPassword)  then
		   CheckUserName=true	
		   sResult= "" 		   
		   
		   Session("Username")=sUserName
		   Session("Password")=sPassword
		   
		   if sRemember="on" then
		   
				dExpireDate = dateserial(year(now)+1, month(now), day(now))

				response.cookies("qo_username") = XORCrypt(sUserName, dbconnect)			
				response.cookies("qo_username").expires = dExpireDate

				response.cookies("qo_password") = XORCrypt(sPassword, dbconnect)			
				response.cookies("qo_password").expires = dExpireDate     		   
		  end if
		  
		  
		  if Trim(ucase(rs.Fields("Privilege")))=ucase("Admin") then
		  
				Session("LoginAdmin") = 1
				Session("f_UserID")=rs.Fields("ID")
		  else
		  
				Session("LoginAdmin") = 0
				Session("f_UserID")=rs.Fields("ID")
		  end if
			
		else
		   sResult = CStr(msg_NoValidLogin)
		   CheckUserName=false
		end if
 else
	sResult = CStr(msg_NoValidLogin)
	CheckUserName=false
 end if
  
  rs.Close
  
end function

Function XORCrypt(xstring,key)
XORCrypt=xstring
'Dim S
'Dim C
'Dim i
	'S=""
	'For i=1 to Len(xString)
		'C=Chr((Asc(Mid(xstring,i,1)) Xor Asc(Mid(key,i,1))) Mod 256)		
		'S=S&c
	'Next
	'XORCrypt=S
End Function

function ValidateNumeric(byval iInteger)
' Description: If iInteger is numeric, then it is just returned.  Otherwise, a zero is returned.

    dim iResult
    if not(isnumeric(iInteger)) or isNull(iInteger) or isEmpty(iInteger) then
       iResult = 0
    else
       iResult = iInteger
    end if
    
    ValidateNumeric = iResult

end function 

















%>