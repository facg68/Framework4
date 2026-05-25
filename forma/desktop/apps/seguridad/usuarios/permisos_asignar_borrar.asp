<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
			Dim conexion, usuIndice, usuMenu, sistema, proceso, usuario, SQLString, menubar, ordenadoPor

            set conexion = Server.CreateObject("ADODB.Connection")
            conexion.open Application("Conn")

            sistema = Request.QueryString("s")   
            proceso = Request.QueryString("p")   
            usuario = Request.QueryString("u")   
            ordenadoPor = Request.QueryString("o")   
        %>
    </head>

    <body>
        <%
            sqlString = "DELETE FROM seg_ProcesosUsuarios " & _
                         "WHERE (CodigoSis = '" & sistema & "') " & _
                         "AND (CodigoProc = '" & proceso & "') " & _
                         "AND (CodigoUsu = '" & usuario & "');"

            conexion.execute(sqlString)    
            conexion.close: set conexion = nothing

            Response.redirect "permisos_asignar.asp?o=" & ordenadoPor & "&s=" & Sistema & "&p=" & Proceso
        %>
    </body>
</html>