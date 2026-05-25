<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
			Dim conexion, usuIndice, usuMenu, usuario, SQLString
            Dim menubar, ordenadoPor, unico, sistema

            set conexion = Server.CreateObject("ADODB.Connection")
            conexion.open Application("Conn")

            rol = Request.QueryString("r")   
            usuario = Request.QueryString("u")   
            estado = Request.QueryString("e")   
            ordenadoPor = Request.QueryString("o")  
            unico = Request.QueryString("un")   	        
            sistema = Request.QueryString("s")                		        
        %>
    </head>

    <body>
        <%
            if estado = 0 then 
                estado = 1
            else
                estado = 0
            end if

            sqlString = "UPDATE seg_RolesUsuarios " & _
                           "SET Activo = " & estado & _
                        " WHERE (CodigoRol = '" & Rol & "') " & _
                           "AND (CodigoUsuario = '" & usuario & "');"

            conexion.execute(sqlString)    
            conexion.close: set conexion = nothing

            Response.redirect "roles_asignar.asp?r=" & rol & "&o=" & ordenadoPor & "&un=" & unico & "&s=" & sistema
        %>
    </body>
</html>