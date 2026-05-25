<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
			<meta http-equiv="X-UA-Compatible" content="IE=edge">
			<title>Polla Fifa World Cup Brasil 2014</title>
			<meta name="description" content="">
			<meta name="viewport" content="width=device-width, initial-scale=1">
			
			<style>
				body {
					background-color:#9F9F9F;
					background-image:url(Imagenes/fondo.jpg);
					text-align:center;
				}
				
				td {
					font-family:Verdana, Arial, Helvetica, sans-serif;
					font-size:12px;
				}
				
				#content h2 {
					text-align: left;
				}
				
				.fondo0 {
					background-color:#F2F2F2;
				}
				
				.fondo1 {
					background-color:#F8FFF0;
				}
			</style>
			
			<script type="text/javascript">
				function VerPolla(Codigo) {
					window.location = "ver.asp?t=" + Codigo
				}
			</script>
		</head>
	
    <body>
			<table width="1000px" border="0" align="center" cellpadding="0" cellspacing="0" style="background-color:#F8F8F8;">
				<tr style="background-color:#000000; color:#FFFFFF;"><td colspan="5"><img src="Imagenes/header.jpg" border="0"></td></tr>
			</table>
			
			<table width="1000px" border="0" align="center" cellpadding="10" cellspacing="0" style="background-color:#F8F8F8;">				
				<tr style="background-color:#000000; color:#FFFFFF;">
					<td width="925px" style="text-align:center; font-size:18px; font-weight:bold;">
						Lista de Apuestas De Usuario
					</td>
					
					<td width="75px" align="center">
						<a href="default.asp">
							<img src="Imagenes/home.png" width="40px" height="40px" border="0" />
						</a>
					</td>
				</tr>
			</table>
			
			<table width="1000px" border="0" align="center" cellpadding="5" cellspacing="0" style="background-color:#F8F8F8;">	
				<!-- 
					Aqui van los campos del formulario... 
				-->
				<%
					dim k, campo, con, sqlString, t, indicefondo, multiplicador, cedula
					
					cedula = Request.Form("txtCedula")
					
					sqlString = "dbo.mundial_ApuestasUsuario '" & Cedula & "';"
					
					set con = Server.CreateObject("ADODB.Connection")
					con.open Application("Conn")				
					set t = con.execute(sqlString)
					
					if (t.bof or t.eof) then
					%>
						<tr>
							<td colspan="9" align="center">
								No encuentro apuestas para esta cedula!
							</td>
						</tr>
					<%
					else
						indicefondo = 1
						multiplicador = -1
						
					%>
						<tr style="background-color:#666666; font-size:12px; color:#FFFFFF; text-align:center;">
							<td>Polla</td>
							<td>Nombre</td>
							<td>Ext.</td>
							<td>Depto.</td>
							<td>F. Confeccion</td>
							<td>Ganador</td>
							<td>Puntos</td>
							<td>Estado</td>
							<td>Accion</td>
						</tr>
					<%
						
						do
							indicefondo = multiplicador + indiceFondo 
							multiplicador = -1 * multiplicador
							
					%>
						<tr class="fondo<%= indicefondo %>" style="color:#<%
																			if t("EnJuego") = 1 then
																				response.write "000000"
																			else
																				response.write "B1B1B1"
																			end if
																		  %>;">
							<td align="center"><%= t("Secuencia") %></td>
							<td><%= t("Nombre") %></td>
							<td><%= t("Telefono") %></td>
							<td><%= t("Departamento") %></td>
							<td><%= t("FechaConfeccion") %></td>
							
							<td align="center">
								<%
									response.write "<img src='"
									
									if t("EnJuego") = 1 then
										response.write "imagenes/banderas/"
									else
										response.write "imagenes/banderas2/"
									end if
									
									response.write t("ImagenGanador")
									response.write "' border='0'  width='33' height='22'>"
									
									response.write "<br />"
									
									response.write "<span style='font-size:9px; color:"
									
									if t("EnJuego") = 1 then
										response.write "#000000"
									else
										response.write "#B1B1B1"
									end if
									
									response.write ";'>" & t("NombreGanador") & "</span>"
								%>
							</td>
							
							<td><strong><%= t("Puntaje") %></strong></td>
							
							<td>
								<img src="Imagenes/estatus<%= t("Estatus") %>.png" border="0" title="<% 
									if t("Estatus") = 0 then 
										response.write "Polla Sin Activar" 
									else 
										response.write "Apuesta en Juego!" 
									end if
								%>">
							</td>
							
							<td style="vertical-align:middle;">
								<input name="" type="button" value="Ver Detalle" onClick="VerPolla('<%= t("Secuencia") %>')">							
							</td>
						</tr>
					<%
							t.movenext
						loop until t.eof
					end if
					
					t.close: set t=nothing
					con.close: set con = nothing
				%>
			</table>
    </body>
</html>