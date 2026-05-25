<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
			Dim conexion, pantalla, usuario

            set conexion = Server.CreateObject("ADODB.Connection")
            conexion.open Application("Conn")

            pantalla = Request.QueryString("p")   
            usuario = Request.QueryString("u")   
        %>
    </head>

    <body>
        <%
            sqlString = "DELETE FROM seg_Anuncios_Pantallas_Usuarios " & _
                         "WHERE (Pantalla = '" & pantalla & "') " & _
                         "AND (Usuario = '" & usuario & "');"


            conexion.execute(sqlString)    
            conexion.close: set conexion = nothing

            Response.redirect "pantallas_asignar.asp?p=" & pantalla
        %>
    </body>
</html>