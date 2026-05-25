<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
	<head>
  		<!-- #include virtual = "/core/includes/kernel/head.inc" -->  	
		<%
			function TieneTransacciones(Usuario, Presupuesto)
				dim cc, tt, sqlString
		
				sqlString = "SELECT ISNULL(COUNT(*), 0) AS Cuantos " & _
								  "FROM pre_Presupuesto_Detalles " & _
								 "WHERE (Usuario = '" & Usuario & "') " & _
								   "AND (Presupuesto = '" & Presupuesto & "') " & _
								   "AND (Aplicado = 1);"
				
				set cc = Server.CreateObject("ADODB.Connection")
				cc.open Application("Conn")
					set tt = cc.Execute(sqlString)
						TieneTransacciones = tt("Cuantos")				
					tt.close: set tt = nothing
				cc.close: set cc = nothing
			end function
		%>		
	</head>
	
    <body plantilla="normal" reserva="165">
  		<!-- #include virtual = "/core/includes/kernel/body.inc" -->  
		<%
			dim usu, pre, c, sqlString

			usu = Request.Cookies("Usuario")
			pre = Request.QueryString("p")
			tipo = Request.QueryString("t")
			estatus = Request.QueryString("y")
			ordenadoPor = Request.QueryString("o")

			if TieneTransacciones(usu, pre) = 0 then
				set c = server.CreateObject("ADODB.Connection")
				c.open Application("Conn")

				'
				'  Borramos los detalles
				'
				sqlString = "DELETE FROM pre_Presupuesto_Detalles " & _
							"WHERE Usuario = '" & usu & "' " & _
								"AND Presupuesto = '" & pre & "';"

				c.execute (sqlString)


				'
				'  Borramos el encabezado
				'
				sqlString = "DELETE FROM pre_Presupuesto_Encabezado " & _
							"WHERE Usuario = '" & usu & "' " & _
								"AND Presupuesto = '" & pre & "';"

				c.execute (sqlString)

				c.close: set c = nothing
				response.redirect "../lista.asp"
			else 
				%>
        			<br />				

					<div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
						<div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
							Borrar Presupuesto <%= pre %>
						</div>
						
						<div style="flex: 0 0 50%; text-align: right;">
							<button type="button" 
									class="form-btn azul" 
								 	style="width: 150px; font-size: 16px; color:white;"
									onclick="volver()">Volver</button>
						</div>
					</div>        

					<div class="main main-scroll">
						<br />

						<div style="font-size: 24px;">
							Error:

							<br /><br />

							Este presupuesto no se puede borrar porque ya tiene (por lo menos) 
							una transaccion aplicada que afecta las cuentas.

							<br /><br />

							Si no se han cerrado las cuentas (en el módulo de cuentas), 
							aún puede editar las transacciones para quitar el estatus de 
							"aplicado", de esa forma podrá eliminar éste presupuesto.

							<br /><br />

							Fin de Proceso
						</div>

						<br />
					</div>
						
					<br/><br />
				<%
			end if
		%>

		<script>
			function volver() {
          		var vinculo = "pre_det_editar.asp?p=<%= pre %>&t=<%= tipo %>&e=<%= estatus %>&o=<%= ordenadoPor %>";			
				window.location.href = vinculo;
			}		
		</script>

		<!-- #include virtual = "/core/includes/kernel/close.inc" -->		
	</body>
</html>