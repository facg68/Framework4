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
		
		<script type="text/javascript">
			function Back() {
				document.Regresar.submit();
			}		
		</script>
	</head>
	
	<body onLoad="Back();">
		<%
			Dim con, sqlString
		
			sqlString = "DELETE FROM mundial_ApuestaS_Enc WHERE Secuencia = '" & request.QueryString("t") & "'"

			if (Usuario_Valido() = 0) then
				response.Redirect "mundial.asp"
			else				
				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				con.execute(sqlString)
				con.close: set con=nothing	
			end if	
		%>

		<div style="display:none;">
			<form id="Regresar" name="Regresar" method="post" action="activar.asp">			
				<input name="txtUsuarioValidado" type="hidden" value="1" />
			</form>	
		</div>
	</body>
</html>
