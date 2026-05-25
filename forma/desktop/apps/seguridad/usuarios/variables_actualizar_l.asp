<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
			Dim conexion, parametro, usuario, SQLString, tipo, valor, ordenadoPor

            set conexion = Server.CreateObject("ADODB.Connection")
            conexion.open Application("Conn")

            tipo = Request.QueryString("t")   

            cadena = Request.QueryString("p")   
            usuario = Request.QueryString("u") 
            valor = Request.QueryString("v")                 
            ordenadoPor = Request.QueryString("o")   

            donde = InStr(cadena, "__")
            sistema = (left(cadena, (donde - 1)))
            parametro = right(cadena, (len(trim(cadena)) - (donde + 1)))                 

            if tipo = "6" then
                valor = replace(valor, "^", "#")
            end if
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

            Response.redirect "editar_usuario_variables.asp?o=" & ordenadoPor & "&u=" & usuario
        %>
    </body>
</html>