<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
			Function Usuario_Valido()
				dim con, t, sqlString
				
				Usuario_Valido = 0
				sqlString = "exec seg_pa_VerificarPermisoUsuario '" & Request.Cookies("Usuario") & "', 'mundial', 'mundial.050'"

				set con = Server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)

				if t("Acceso") = 1 then
					Usuario_Valido = 1
				end if
				
				t.close: set t=nothing
				con.close: set con=nothing			
			End Function	            
        %>    
    </head>

    <body>
        <%
            dim key, cc

			if (Usuario_Valido() = 1) then
                key = Request.Cookies("mundial_key")

                if key = "dslkwrtywbfsjbvwegowuienweixmlkjri" then
                    set cc = Server.CreateObject("ADODB.Connection")
                    cc.open Application("Conn")

                    '
                    ' Cerramos el periodo de creacion de pollas
                    '

                    cc.execute("UPDATE mundial_Estatus SET Estatus = 0 WHERE Codigo = 'Activar';")

                    cc.close: set cc = nothing
                end if
            end if

            response.redirect "mundial.asp"
        %>
    </body>
</html>