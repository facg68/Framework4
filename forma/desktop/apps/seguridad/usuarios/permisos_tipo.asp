<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
			Dim conexion, usuIndice, usuMenu, sistema, proceso, usuario, estatus, SQLString, menubar, ordenadoPor

            set conexion = Server.CreateObject("ADODB.Connection")
            conexion.open Application("Conn")

            sistema = Request.QueryString("s")   
            proceso = Request.QueryString("p")   
            usuario = Request.QueryString("u")   
            estatus = Request.QueryString("e")   
            ordenadoPor = Request.QueryString("o")   
        %>
    </head>

    <body>
        <%
            sqlString = "UPDATE seg_ProcesosUsuarios " & _
                           "SET TipoProceso = " & estatus & _
                         "WHERE (CodigoSis = '" & sistema & "') " & _
                           "AND (CodigoProc = '" & proceso & "') " & _
                           "AND (CodigoUsu = '" & usuario & "');"

            conexion.execute(sqlString)    
            conexion.close: set conexion = nothing

            Response.redirect "permisos_asignar.asp?o=" & ordenadoPor & "&s=" & Sistema & "&p=" & Proceso
        %>
    </body>
</html>