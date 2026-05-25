<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
			Dim conexion, usuIndice, usuMenu, tabla, usuario, password, SQLString, mantener, menubar, tt, ordenadoPor
            Dim unico, sistema

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
            sqlString = "INSERT INTO seg_RolesUsuarios(CodigoRol, CodigoUsuario, Activo) " & _
                        "VALUES('" & Rol & "', '" & usuario & "', 1);"

            conexion.execute(sqlString)    
            conexion.close: set conexion = nothing

            Response.redirect "roles_asignar.asp?r=" & rol & "&o=" & ordenadoPor & "&un=" & unico & "&s=" & sistema
        %>
    </body>
</html>