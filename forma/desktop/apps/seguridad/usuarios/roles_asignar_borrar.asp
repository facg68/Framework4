<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
			Dim conexion, usuIndice, usuMenu, tabla, usuario, password, SQLString, mantener
            Dim menubar, tt, ordenadoPor, unico, sistema

            set conexion = Server.CreateObject("ADODB.Connection")
            conexion.open Application("Conn")

            rol = Request.QueryString("r")   
            usuario = Request.QueryString("u")   
            ordenadoPor = Request.QueryString("o")   
            unico = Request.QueryString("un")   	        
            sistema = Request.QueryString("s")               	        
        %>
    </head>

    <body>
        <%
            sqlString = "DELETE FROM seg_RolesUsuarios " & _
                        "WHERE CodigoUsuario = '" & usuario & "' " & _
                        "AND CodigoRol = '" & rol & "';"

            conexion.execute(sqlString)  
            conexion.close: set conexion = nothing

            Response.redirect "roles_asignar.asp?r=" & rol & "&o=" & ordenadoPor & "&un=" & unico & "&s=" & sistema
        %>
    </body>
</html>