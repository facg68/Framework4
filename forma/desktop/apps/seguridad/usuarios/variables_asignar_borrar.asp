<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
			Dim conexion, usuIndice, usuMenu, sistema, parametro, usuario, SQLString, menubar, ordenadoPor, tt

            set conexion = Server.CreateObject("ADODB.Connection")
            conexion.open Application("Conn")

            sistema = Request.QueryString("s")   
            parametro = Request.QueryString("p")   
            usuario = Request.QueryString("u")   
            ordenadoPor = Request.QueryString("o")   

            function ValorDefault(Sistema, Parametro)
                set tt = conexion.execute("SELECT ValorDefault FROM seg_Parametros WHERE (Sistema= '" & sistema & "') AND (Parametro = '" & parametro & "');")
                    ValorDefault = tt("ValorDefault")
                tt.close: set tt = nothing
            end function
        %>
    </head>

    <body>
        <%
            sqlString = "DELETE FROM seg_Usuarios_Parametros " & _
                         "WHERE (Usuario = '" & usuario & "') " & _
                           "AND (Sistema = '" & sistema & "') " & _
                           "AND (Parametro = '" & parametro & "');"

            conexion.execute(sqlString)    
            conexion.close: set conexion = nothing

            Response.redirect "variables_asignar.asp?o=" & ordenadoPor & "&s=" & Sistema & "&p=" & parametro
        %>
    </body>
</html>