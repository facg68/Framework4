<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
	<head>
		<%
			Function Usuario_Valido()
				dim con, t, sqlString
				
				Usuario_Valido = 0
				sqlString = "exec seg_pa_VerificarPermisoUsuario '" & Request.Cookies("Usuario") & "', 'mundial', 'mundial.040'"

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
			dim Codigo, con, sqlString

			if (Usuario_Valido() = 0) then
				response.Redirect "mundial.asp"
			else			
				Codigo = Request.Form("txtCodigoUnico")
				
				sqlString = "UPDATE mundial_Apuestas_Enc SET Estatus = 1 WHERE Secuencia = '" & Codigo & "';"
				
				set con = Server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				con.execute(sqlString)
				con.close: set con = nothing
				
				Response.Redirect "boleto.asp?t=" & Codigo
			end if
		%>
	</body>
</html>