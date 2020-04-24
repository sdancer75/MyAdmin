<%@CodePage=1253 Language=VBScript%>

<!-- #include file="Greek.asp" -->
<!-- #include file="../functions.asp" -->

<%
  Dim objConn, strDBPath, strDomain, smtpServer, strLink
  Dim fromAddr, mailComp, strSiteTitle
  
  'Change to the title of your site IE: HTMLJunction
  strSiteTitle = "Paradox Interactive Multimedia Software"
  'The address to your website WITHOUT "http://"
  strDomain = "www.paradoxinteractive.gr"
  





  'Your mail server
  smtpServer = "mail.paradoxinteractive.gr"
  'your email address on the mail server
  fromAddr = "info@paradoxinteractive.gr"
  
  'Select your email component
  mailComp = "CDOSYS"
  'mailComp = "CDONTS"
  'mailComp = "JMail"
  'mailComp = "ASPMail"

  
  
			   
			   

%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=windows-1253"> 

<html><head><title>NewsLetter</title>
<style>
			P { text-align: justify; }
</style>
</head>

<body>
<%

  	SQL = "SELECT * FROM newsLetter WHERE confirm = 'yes';"
	rs.CursorLocation = 3	'adUseClient
	rs.CursorType = 3		'adOpenStatic
	rs.ActiveConnection = MyConn
  	
	rs.Open SQL
	if rs.EOF = false then
		rs.MoveLast
		strCount = rs.RecordCount 
		rs.close
	end if

  
  If Request.Form("action") = "" Then
%>


<div align="left">    
  <form action="SendNewsLetter.asp" method="post">
  <input type="hidden" name="purge" value="yes">
  <%=PurgeDB%>&nbsp;&nbsp;<input type="submit" value="Yes">
  </form>
  <br><br>
  <span style="font-family:arial;font-size:12px;color:#000080;font-weight:bold">
    <%=MembersPart1%> <%= strCount %> <%=MembersPart2%>
  </span>
  <br /><br />
  <form action="SendNewsLetter.asp" method="post">
  <input type="hidden" name="action" value="send" />
  
    <table width="50%">
	  <tr>
	    <td><span class="thrd"><%=SubjectMsg%>:</span>&nbsp;&nbsp;<input type="text" name="nwsubject" size="30" value="Newsletter" /></td>
	  </tr>
	  <tr>
	    <td>
		  <textarea name="msg" cols="50" rows="10"></textarea>
		</td> 
	  </tr>
	  <tr>
	    <td align="center"><input type="radio" name="version" value="html" checked />&nbsp;<span style="font-family:arial;font-size:12px;color:#000080;font-weight:bold">HTML&nbsp;&nbsp;<input type="radio" name="version" value="text" />&nbsp;TEXT</span></td>
	  </tr>
	  <tr>
	    <td align="center"><input type="submit" value="<%=send%>"></td>
	  </tr>
	</table>
  </form>
</div>
<%
  Else
    Dim mailObj, cdoMessage, cdoConfig, addrList, strEmailMsg, subject
	Dim strEmail
	
	
	'Reset all views
	Sql = "Update Newsletter set FeedBackView = 0"
	MyConn.execute(sql)
	
	strFeedbackView = "<script language=javascript> </script>"
	
	SQL = "SELECT * FROM newsLetter WHERE confirm = 'yes';"
	rs.Open SQL, MyConn, adOpenForwardOnly, adLockOptimistic, adCmdText
	strCount = 0
	If Not rs.EOF Then
	  Do While Not rs.EOF

		  
			'Customize the footer that gets inserted at the bottom of the newsletter
			If Request.Form("version") = "html" Then 'html link
			  strLink = "<a href=""http://" & strDomain & "/gr/FeedBack.asp?Type=NewsLetter&Cmd=Verify&email=" & Request.form("newsletter_email") & "&cancel=yes>" & _
						  "http://" & strDomain & "/gr/FeedBack.asp?Type=NewsLetter&Cmd=Verify&email=" & Request.form("newsletter_email") & "&cancel=yes>"
			Else ' text link
			  strLink = "http://" & strDomain & "/gr/FeedBack.asp?Type=NewsLetter&Cmd=Verify&email=" & Request.form("newsletter_email") & "&cancel=yes>"
			End If
  
	
			strEmail = rs("email")
			subject = request.form("nwsubject")
		  'send email so subscriber can confirm
			strEmailMsg = Replace(Request.Form("msg"),Chr(13) & Chr(10),"<br>")
			If Request.Form("version") = "html" Then
			  strEmailMsg = strEmailMsg & "<br><br><hr widht=""100%""> <span style=""font-family:arial;font-size:10px;color:#4682b4;"">" & strFooter & "<br>" & strLink & "<br><br>" & strSiteTitle & "</span>"
			Else
			  strEmailMsg = Replace(strEmailMsg,"<br>",Chr(13) & Chr(10))
			  strEmailMsg = strEmailMsg & Chr(13) & Chr(10)& Chr(13) & Chr(10) & "_______________________________________________" & Chr(13) & Chr(10) & strFooter & Chr(13) & Chr(10) & strLink & Chr(13) & Chr(10) & strSiteTitle
			  
			End If
			
		  'Add Feedback script code
		  
		  strEmailMsg = strEmailMsg & strFeedbackView
		  'Send email based on mail component.
	
			'Send email (CDONTS version). Note: CDONTS doesn't support a reply-to
		  'address and has no error checking.
	
			if mailComp = "CDONTS" then
			  set mailObj = Server.CreateObject("CDONTS.NewMail")
			  If Request.Form("version") = "html" Then
			    mailObj.BodyFormat = 0
			    mailObj.MailFormat = 0
			  Else
			    mailObj.BodyFormat = 1
			  End If
			  mailObj.From = fromAddr
			  mailObj.To = strEmail
			  mailObj.Subject = subject
			  mailObj.Body = strEmailMsg
			  mailObj.Send
			  set mailObj = Nothing
			end if
	
			'Send email (CDOSYS version).
	
			if mailComp = "CDOSYS" then
			  set cdoMessage = Server.CreateObject("CDO.Message")
			  set cdoConfig = Server.CreateObject("CDO.Configuration")
			  cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			  cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpServer
			  cdoConfig.Fields.Update
			  set cdoMessage.Configuration = cdoConfig
			  cdoMessage.BodyPart.charset = "windows-1253"
			  cdoMessage.From =  fromAddr
			  cdoMessage.To = strEmail
			  cdoMessage.Subject = subject
			  If Request.Form("version") = "html" Then
			    cdoMessage.HtmlBody = strEmailMsg
			  Else
			    cdoMessage.TextBody = strEmailMsg
			  End If
			  on error resume next
			  cdoMessage.Send
			  if Err.Number <> 0 then
				SendMail = EmailFail & Err.Description & "."
			  end if
			  set cdoMessage = Nothing
			  set cdoConfig = Nothing
			end if
	
			'Send email (JMail version).
	
			if mailComp = "JMail" then
			  set mailObj = Server.CreateObject("JMail.SMTPMail")
			  mailObj.Silent = true
		    mailObj.Charset = "windows-1253" 'Ελληνικοί χαρακτήρες
		    mailObj.ISOEncodeHeaders = false            
			  mailObj.ServerAddress = smtpServer
			  mailObj.Sender = fromAddr
			  mailObj.ReplyTo = fromAddr
			  mailObj.Subject = subject
			  addrList = Split(strEmail, ",")
			  for each addr in addrList
		      mailObj.AddRecipient Trim(addr)
		    next
			  If Request.Form("version") = "html" Then mailObj.ContentType = "text/html"
			  mailObj.Body = strEmailMsg
			  if not mailObj.Execute then
				SendMail = EmailFail & mailObj.ErrorMessage & "."
			  end if
			end if
	
			'Send email (ASPMail version).
	
			if mailComp = "ASPMail" then
			  set mailObj = Server.CreateObject("SMTPsvg.Mailer")
			  mailObj.FromAddress = fromAddr
			  mailObj.RemoteHost  = smtpServer
			  mailObj.ReplyTo = fromAddr
			  for each addr in Split(strEmail, ",")
		      mailObj.AddRecipient "", Trim(addr)
		    next
			  mailObj.Subject = subject
			  If Request.Form("version") = "html" Then mailObj.ContentType = "text/html"
			  mailObj.BodyText = strEmailMsg
			  if not mailObj.SendMail then
				SendMail = EmailFail & mailObj.Response & "."
			  end if
		  end if
		  
		  

			    
			If Err.Number <> 0 then
			  Response.Write strEmail & " -> " & SendMail & vbCrLf
			Else
			  Response.Write "<div align=""center""><br><br>" & _
			                 "<span style=""font-family:arial;font-size:12px;color:#000080;font-weight:bold;text-align:left"">" & _
			                 strEmail & " -> " & EmailSuccess & "</span></div><br>" 
			End If

			rs.MoveNext  
		Loop
		rs.Close	  
	End If
		
	
  End If

  
  If Request.Form("purge") = "yes" Then      
      sql = "SELECT * FROM newsLetter WHERE confirm = 'no';"
      rs.Open sql, MyConn, adOpenForwardOnly, adLockOptimistic, adCmdText
      If Not rs.EOF Then
          Do While Not rs.EOF
		    If DateDiff("y",rs("Date"),Now) > 7 then
              rs.Delete
			End If
	        rs.MoveNext
	      Loop
      End If
      rs.Close
     
  End If

 

%>
</body>
</html>

