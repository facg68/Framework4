<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
			Dim conexion, sistema, parametro, usuario, SQLString, valor, ordenadoPor

            set conexion = Server.CreateObject("ADODB.Connection")
            conexion.open Application("Conn")

            sistema = Request.QueryString("s")   
            parametro = Request.QueryString("p")   
            usuario = Request.QueryString("u")   
            valor = Request.QueryString("v")   
            ordenadoPor = Request.QueryString("o")   
        %>
    </head>

    <body>
        <%
            sqlString = "UPDATE seg_Usuarios_Parametros " & _
                           "SET Valor = '" & valor & "' " & _
                         "WHERE (Usuario = '" & usuario & "') " & _
                           "AND (Sistema = '" & sistema & "') " & _
                           "AND (Parametro = '" & parametro & "');"

            conexion.execute(sqlString)    
            conexion.close: set conexion = nothing

            Response.redirect "variables_asignar.asp?o=" & ordenadoPor & "&s=" & Sistema & "&p=" & parametro
        %>
    </body>
</html>