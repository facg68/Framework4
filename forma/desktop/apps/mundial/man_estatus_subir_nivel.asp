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

			function Etapa()
				sqlString = "SELECT Estatus FROM dbo.mundial_Estatus WHERE Codigo = 'Etapa';"
							
				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)
				
				if not (t.eof or t.bof) then
					Etapa = t("Estatus")
				end if
			
				t.close: set t=nothing
				con.close: set con=nothing
			end function	                       
        %>    
    </head>

    <body>
        <%
            dim key, cc, e

			if (Usuario_Valido() = 1) then
                key = Request.Cookies("mundial_key")

                if key = "dslkwrtywbfsjbvwegowuienweixmlkjri" then
                    set cc = Server.CreateObject("ADODB.Connection")
                    cc.open Application("Conn")

                    '
                    ' Subimos el Nivel de la Etapa
                    '
                    e = Etapa() + 1
                    if e > 6 then  e = 0

                    cc.execute("UPDATE mundial_Estatus SET Estatus = " & e & " WHERE Codigo = 'Etapa';")

                    cc.close: set cc = nothing
                end if
            end if

            response.redirect "mundial.asp"
        %>
    </body>
</html>
