<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Polla Mundial</title>
		<!-- #include virtual = "/core/includes/kernel/head.inc" -->
		
        <style>
            body {
                background-color:#9F9F9F;
                background-image:url(Imagenes/fondo.jpg);
            }
			
            td {
                font-family:Verdana, Arial, Helvetica, sans-serif;
                font-size:12px;
				padding: 0px; 
				border-spacing; 0px; 
				border-style: none;"				
            }
			
            #content h2 {
                text-align: left;
            }
			
			.fondo0 {
				background-color:rgb(235, 243, 252);
			}
			
			.fondo1 {
				background-color:rgb(255,255,255);
			}

			tr:not(:last-child) { border: none !important; }
        </style>
    </head>
	
    <body>
		<!-- #include virtual = "/core/includes/kernel/body.inc" -->

		<br />

		<div class="main" style="background-color: rgba(255, 255, 255, 0.25);">
			<div class="line">
				<table width="100%" style="margin-left: auto; margin-right: auto; padding: 0px; border-style: none; border-spacing; 0px; background-color:#000000;">
					<tr>
						<td><img src="Imagenes/header.jpg" style="border-style: none; width: 100%;" /></td>
					</tr>

					<tr>
						<td>
							<table width="100%" style="margin-left: auto; margin-right: auto; padding: 0px; border-spacing; 0px;">
								<tr>
									<td width="95%" 
										style="font-family: Verdana, Arial, Helvetica, sans-serif; 
											   font-size: 18px; text-align: center middle; 
											   color:rgb(255, 255, 255);">
										&nbsp;&nbsp;LISTA DE APUESTAS DEL USUARIO
									</td>
									
									<td width="5%" style="text-align: right;">
										<a href="default.asp">
											<img src="Imagenes/home.png" width="50px" height="50px" />
										</a>
									</td>
								</tr>
							</table>				
						</td>
					</tr>

					<tr>
						<td>
							<table width="100%" style="margin-left: auto; margin-right: auto; padding: 0px; border-spacing: 0px; border-style: none;">
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
												<td colspan="9" style="text-align: center;">
													No encuentro apuestas para esta cedula!
												</td>
											</tr>
										<%
									else
										indicefondo = 1
										multiplicador = -1			
										%>
											<tr style="background-color:#666666; font-size:12px; color:#FFFFFF; text-align: center;">
												<td style="padding: 8px; width: 15%">Polla</td>
												<td style="padding: 8px; width: 20%">Nombre</td>
												<td style="padding: 8px; width: 10%">Ext.</td>
												<td style="padding: 8px; width: 15%">Depto.</td>
												<td style="padding: 8px; width: 15%">Fecha</td>
												<td style="padding: 8px; width: 10%">Ganador</td>
												<td style="padding: 8px; width:  5%">Puntos</td>
												<td style="padding: 8px; width:  5%">Estado</td>
												<td style="padding: 8px; width:  5%">Accion</td>
											</tr>
										<%			
											do
												indicefondo = multiplicador + indiceFondo 
												multiplicador = -1 * multiplicador					
										%>
											<tr class="fondo<%= indicefondo %>" 
												style="color:#<%
																if t("EnJuego") = 1 then
																	response.write "000000"
																else
																	response.write "B1B1B1"
																end if
															%>;">
											<td style="text-align: center; font-weight: bold;"><%= t("Secuencia") %></td>
											<td style="padding: 5px;"><%= t("Nombre") %></td>
											<td style="padding: 5px;"><%= t("Telefono") %></td>
											<td style="padding: 5px;"><%= t("Departamento") %></td>
											<td style="padding: 5px;"><%= t("FechaConfeccion") %></td>
											
											<td style="text-align: center; padding: 5px;">
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
											
											<td style="padding: 5px;"><strong><%= t("Puntaje") %></strong></td>
											
											<td style="padding: 5px;">
												<img src="Imagenes/estatus<%= t("Estatus") %>.png" style="border-style: none;" title="<% 
													if t("Estatus") = 0 then 
														response.write "Polla Sin Activar" 
													else 
														response.write "Apuesta en Juego!" 
													end if
												%>">
											</td>
											
											<td style="vertical-align:middle; padding: 5px;">
												<button class="form-btn verde tiny" type="button" onClick="VerPolla('<%= t("Secuencia") %>')">
													<i class='fa fa-eye fa-normal' title='Borrar lista'></i>
												</button>
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
						</td>
					</tr>
				</table>
			</div>
		</div>

		<script type="text/javascript">
			function VerPolla(Codigo) {
				window.location = "ver.asp?t=" + Codigo
			}
		</script>	
		<!-- #include virtual = "/core/includes/kernel/close.inc" -->		
    </body>
</html>