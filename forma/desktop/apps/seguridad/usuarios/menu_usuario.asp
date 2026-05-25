<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
		<!-- #include virtual = "/core/includes/menu_pass.inc" -->
		
        <%
			Dim usuario, ordenadoPor

            usuario = Request.QueryString("u")   
            ordenadoPor = Request.QueryString("o")   
        %>
    </head>

    <body>
        <%
            CrearMenu usuario
            ActualizarSnippetsUsuario usuario
            ResetearHomePage usuario
            Response.redirect "lista.asp?o=" & ordenadoPor 
        %>
    </body>
</html>