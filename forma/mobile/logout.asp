<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
  <head>
    	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
		<!-- #include virtual = "/core/includes/no_sql_injection.asp" -->
	</head>
	
	<body>
		<%
			Response.Cookies("usuario").Expires = Date() - 1
			Response.Cookies("nombre").Expires = Date() - 1		
			Response.Redirect("/pc_mobile.asp")
		%>	
	</body>
</html>