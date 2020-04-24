<%@codepage=1253 language = "VBScript" %>
<% 


    option explicit 

%>

<!--#include file ="include/adovbs.inc"-->
<!--#include file ="include/ute_definition.inc"-->

<%

	SetLocale(1032)
	session.lcid=2057
	Session.Timeout=30
	
    Dim sDSN,dbconnect
    ' To use a DSN-Less Connection use the following sDSN definition.
    ' !!! By using this, UTE is able to detect Primary Keys accurately.
    dbconnect = "Driver={SQL Server}; Server=" & sMSSQLServerAddress & "; Database=" &  sMSSQLInitialCatalog & "; Uid=" & sMSSQLUsername & "; Pwd=" & sMSSQLPassword
    sDSN = dbconnect
    ' To use a DSN (ODBC) Connection use the following sDSN defintion.
    ' You need to setup an ODBC data source.
    ' !!! By using this, UTE is *NOT* always able to detect Primary Keys accurately.                                        
    ' sDSN = "test"

    Dim ute,test
    Set ute = new clsUTE

    ute.DBName      = "All-About-Bet.com"  ' Name of Database. For display purpose only
    'ute.ReadOnly    = True       ' readonly mode
    ute.ListTables  = true      ' display toolbutton to list all tables within db
    ute.Filters     = true      ' display toolbutton to define and activate filters
    ute.Export      = true      ' display toolbutton to export all data to CSV file
    ute.SQL         = true      ' display toolbutton to show current sql statement
    ute.Definitions = true      ' display toolbutton to show field defintions
    ute.MainMenu	= "admin.asp"
    ute.TableLookUp	= true

    ute.Init sDSN   ' init must be called *before* any HTML code is
                    ' is written, otherwise the export feature doesn't work !
	
	
	if Request.QueryString ("name")="Articles" then
		ute.AddLookUpTable("TipsterID=sql(select TipsterID,NickName from Tipsters where TipsterID=v_Field{TipsterID:NickName}pkvalue{TipsterID})")		
		ute.AddDefaultValue("RegDate=" & Date())
		ute.AddRichEdit("Article")
		ute.AddLookUpTable("ArticleCategory=Lookup(1|Γενικό άρθρο;2|Golden Match;3|Αξίζει να δεις;4|Τι παίζει η Ασία;5|Γερμανία;6|Εκπλήσεις))")
		ute.AddHookBrowse("Photo")
			
	end if
	
	if Request.QueryString ("name")="News" then
		
		ute.AddDefaultValue("Date=" & Date())
		ute.AddDefaultValue("EndDate=" & CDate(Date()+1))

			
	end if	
	
	if Request.QueryString ("name")="Tipsters" then

		ute.AddRichEdit("CV")
		ute.AddHookBrowse("Photo")
		ute.AddHTMLCode("TipsterAdmin= Τα κείμενα του TipsterAdmin εμφανίζονται πρώτα για την ίδια ημερομηνία")
			
	end if
	
	
	if Request.QueryString ("name")="BannersLeftSide" then

		ute.AddHookBrowse("BannerLeftTitle")
		ute.AddHTMLCode("Width=Το μέγιστο επιτρεπόμενο πλάτος είναι 200 pixels")
			
	end if	
	
	if Request.QueryString ("name")="BannersRightSide" then

		ute.AddHookBrowse("BannersRightSide")
		ute.AddHTMLCode("Width=Το μέγιστο επιτρεπόμενο πλάτος είναι 200 pixels")
			
	end if		
	
	
	
	
	
	if Request.QueryString ("name")="UserFiles" then
		ute.AddLookUpTable("Privilege=Lookup(Admin|Admin)")		
	end if	
	

%>
<!doctype html public "-//W3C//DTD HTML 3.2//EN">
<html>
<head>
  <title><%=ute.HeadLine%> - Universal Table Editor</title>
  <link rel="stylesheet" type="text/css" href="ute_style.css">
  <META HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=windows-1253"> 
</head>
<body bgcolor="#FFFFFF" link="#0000A0" vlink="#0000A0" alink="#0000A0">
<%
    ute.Draw
    Set ute = Nothing
%>
</body>
</html>
