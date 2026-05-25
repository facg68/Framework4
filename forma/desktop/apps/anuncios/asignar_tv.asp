<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim cc, tt, sqlString, pantalla

    pantalla = Request.QueryString("a")

    Response.Cookies("Pantalla") = pantalla
    Response.Cookies("Pantalla").Expires = Date() + 3000

    response.redirect "pantallas.asp"    
%>