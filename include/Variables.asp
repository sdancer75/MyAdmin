<!-- #include file="WebSiteAccount.asp" -->

<%
Dim MyConn,rs,rsMaster
Dim dbconnect,sResult
Dim tbl_UsersTable,f_Username,f_Password,html_username,html_password,html_submit,msg_NoValidLogin
Dim sAuthAccessTitle,sLoginStartPage,sMainMenuTitle


sResult=""
'==================== User Defined Variables ============================



tbl_UsersTable					="Users"
f_Username						="username"
f_Password						="password"
Dim f_RelationField			
Dim f_UserID						


html_username					="Username"
html_password					="Password"
html_submit						="submit"

msg_NoValidLogin				="No Valid Username or Password"
sAuthAccessTitle				=""
sLoginStartPage					="admin.asp"

sMainMenuTitle					="Administrator's page"



f_RelationField					="CSN"
f_UserID=2


%>

