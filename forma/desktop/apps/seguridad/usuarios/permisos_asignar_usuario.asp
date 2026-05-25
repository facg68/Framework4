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
            sqlString = "INSERT INTO seg_ProcesosUsuarios(CodigoSis, CodigoProc, CodigoUsu, Activo, TipoProceso) " & _
                        "VALUES('" & sistema & "', '" & proceso & "', '" & usuario & "', 1, 1);"

            conexion.execute(sqlString)    
            conexion.close: set conexion = nothing

            Response.redirect "permisos_asignar.asp?o=" & ordenadoPor & "&s=" & Sistema & "&p=" & Proceso
        %>
    </body>
</html>