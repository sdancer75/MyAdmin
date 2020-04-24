<%@ Language=VBScript%>



<!-- #include file="include/functions.asp" -->
<%

	if CheckUsername( Session("UserName"),Session("Password"),"off" )=false then
		if CheckUsername( LoadCookie("qo_username"),LoadCookie("qo_password"),"off" )=false then
			Response.Redirect "invalidlogin.asp"	
		end if
	end if

	Response.CacheControl = "no-cache"
	Response.Expires = -1	
	
	
  If UCase(Request.QueryString("Type")) = UCase("ClearOldLogs") Then      
      sql = "SELECT * FROM Logs"
      rs.Open sql, MyConn, adOpenForwardOnly, adLockOptimistic, adCmdText
      If Not rs.EOF Then
          Do While Not rs.EOF
		    If DateDiff("y",rs("EventDate"),Now) > 7 then
              rs.Delete
			End If
	        rs.MoveNext
	      Loop
      End If
      rs.Close
     
  End If
  
  
    If UCase(Request.QueryString("Type")) = UCase("DeleteOldOrders") Then      
      sql = "SELECT * FROM Orders"
      rs.Open sql, MyConn, adOpenForwardOnly, adLockOptimistic, adCmdText
      If Not rs.EOF Then
          Do While Not rs.EOF
		    If DateDiff("y",rs("EventDate"),Now) > 365 then
              rs.Delete
			End If
	        rs.MoveNext
	      Loop
      End If
      rs.Close
     
  End If
  		
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head><title>MyAdmin copyright (c) Paradox Interactive</title>
<link href="include/style.css" rel="stylesheet" type="text/css">
<!-- #include file="include/METATags.asp" -->
</head>

<div align="left">
	<h3><%=sMainMenuTitle%></h3>
	<span class=smalltext>Logged as <b><%=LoadCookie("qo_username")%></span>  <a href="logoff.asp">Log out</a>
	<Hr width=100%>
	<br><br>
	<UL>
				
		<LI><a href="ute.asp?name=Articles"> Articles </a></LI>
		<LI><a href="ute.asp?name=Tipsters"> Tipsters </a></LI>
		<LI><a href="ute.asp?name=Members"> Members </a></LI>
		<LI><a href="ute.asp?name=News"> News </a></LI>		
		<LI><a href="ute.asp?name=Tickers"> Ticker (scrolling news) </a></LI>		
		<LI><a href="ute.asp?name=Links"> Links </a></LI>
		
		<HR width=50%>
		
		<LI><a href="ute.asp?name=BannersLeftSide"> Banners Left Side </a></LI>				
		<LI><a href="ute.asp?name=BannersRightSide"> Banners Right Side </a></LI>		
		

		<HR width=50%>

		<LI><a href="ute.asp?name=Users"> Users Password</a></LI>	
		<LI><a href="ute.asp?name=HitCounter"> HitCounter </a></LI>					
					
		
		
		<br><br>

		

		
		
	</UL>
	
	
</div> 





</body>
</html>


 <script LANGUAGE="javascript" TYPE="text/javascript">

	
	function AreYouSure(type)
	{
		switch (type) {
		
		case 'order':
		
						
				if (confirm('Είσαι σίγουρος ότι θέλεις να διαγράψεις τα Order που είναι παλαιότερα του 1 έτους ?'))				
				{					
					window.location.href = 'Admin.asp?Type=DeleteOldOrders' 
				}
				
				break;
		
		case 'logs':
		
				if (confirm('Είσαι σίγουρος ότι θέλεις να διαγράψεις τα logs που είναι παλαιότερα των 7 ημερών ?'))				
				{					
					window.location.href = 'Admin.asp?Type=ClearOldLogs' 
				}
				
				break;	
		}	
	}
		
</script> 
