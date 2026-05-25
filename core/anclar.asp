<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <%
        dim conn

        set conn = Server.CreateObject("ADODB.Connection")
        conn.open Application("Conn")
    %>

    <head>
        <%
			function Accion(Sistema, Proceso)
				dim t, cmdString

				cmdString = "SELECT proAction " & _
							  "FROM seg_Procesos " & _
							 "WHERE (proSistema = '" & Sistema & "') " & _
                               "AND (proCodigo = '" & Proceso & "');"

				set t = conn.execute(cmdString)
                    Accion = "/forma/desktop/apps/" & Sistema & "/" & t("proAction") & ".asp"
				t.close: set t = nothing
			end function
        %>
    </head>

    <body>
        <%
            dim sis, proc, ope, sqlString, vinculo

            sis = Request.QueryString("s")
            proc = Request.QueryString("p")
            ope = Request.QueryString("a")

            vinculo = Accion(sis, proc)

            if ope = 1 then
                '
                ' Añadir Nuevo Vínculo 
                '
                sqlString = "INSERT INTO seg_Usuarios_Shortcuts(codUsuario, codSistema, codProceso) " & _
                            "VALUES('" & Request.Cookies("Usuario") & "','" & sis & "','" & proc & "');"
            else
                '
                ' Eliminar Vinculo
                '
                sqlString = "DELETE FROM seg_Usuarios_Shortcuts " & _
                             "WHERE (codUsuario = '" & Request.Cookies("Usuario") & "') " & _
                               "AND (codSistema = '" & sis & "') " & _
                               "AND (codProceso = '" & proc & "');"

            end if

            '
            ' Ejecutamos el comando...
            '
            conn.execute(sqlString)

            '
            ' Y volvemos a cargar la pagina...
            '
            response.redirect vinculo

        %>
    </body>

    <%
        conn.close: set conn = nothing
    %>    
</html>