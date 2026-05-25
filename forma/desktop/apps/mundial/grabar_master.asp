<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
	<head>
		<title>Actualizar Registro Master</title>
		<%
			Function Usuario_Valido()
				dim con, t, sqlString
				
				Usuario_Valido = 0
				sqlString = "exec seg_pa_VerificarPermisoUsuario '" & Request.Cookies("Usuario") & "', 'mundial', 'mundial.020'"

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
			dim con, sqlString

			if (Usuario_Valido() = 0) then
				response.Redirect "mundial.asp"
			else			
				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				
				for k = 1 to 31
					if request.Form("Polla" & k) <> "" then
						sqlString = "UPDATE mundial_Master SET CodigoEquipo = '" & request.Form("Polla" & k) & "' WHERE IndiceCuadro = " & k & ";"
						con.execute(sqlString)
					end if
				next

				con.close: set con = nothing

				response.redirect "mundial.asp"
			end if
		%>
	</body>
</html>
