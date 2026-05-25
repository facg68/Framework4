<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
    	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <title>Polla Mundial</title>
        <meta name="description" content="">
        <meta name="viewport" content="width=device-width, initial-scale=1">

        <!-- #include virtual = "/core/includes/kernel/head.inc" -->   
        <%
            thisSystem = "mundial"
            thisProcess = "mundial.030"
            SysLockOut
        %>    

        <style>
            body {
                background-color:#9F9F9F;
                background-image:url(Imagenes/fondo.jpg);
            }
			
            td {
                font-family:Arial, Helvetica, sans-serif;
                font-size:12px;
            }
			
            #content h2 {
               text-align: left;
            }
			
			.fondo0 {
				background-color:rgb(255, 255, 255);
			}
			
			.fondo1 {
				background-color:rgb(217, 231, 255);
			}

			tr:not(:last-child) { border: none !important; }			
        </style>

		<%
			Function Usuario_Valido()
				dim con, t, sqlString
				
				Usuario_Valido = 0
				sqlString = "exec seg_pa_VerificarPermisoUsuario '" & Request.Cookies("Usuario") & "', 'mundial', 'mundial.030'"

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
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->
		<%
			if (Usuario_Valido() = 0) then
				response.Redirect "mundial.asp"
			else			
		%>

				<table style="width: 95%; border-style: none; margin-left: auto; margin-right: auto; padding: 0px; border-spacing: 0px; background-color:#F8F8F8;">
					<tr style="background-color:#000000; color:#FFFFFF;">
						<td>
							<img src="Imagenes/header.jpg" style="border-style: none; width: 100%;">
						</td>
					</tr>
				</table>
				
				<table style="width: 95%; border-style: none; margin-left: auto; margin-right: auto; padding: 10; border-spacing: 0; background-color:#F8F8F8;">
					<tr style="background-color:#000000; color:#FFFFFF;">
						<td width="750%" style="text-align:center; font-size:18px; font-family: arial;">
							Habilitar / Deshabilitar Equipos
						</td>
						
						<td style="width: 25%; text-align: center;">
							<a href="default.asp">
								<img src="Imagenes/home.png" style="width: 40px; height: 40px; border-style: none;" />
							</a>
						</td>
					</tr>
				</table>

				<table style="width: 95%; border-style: none; margin-left: auto; margin-right: auto; padding: 5; border-spacing: 0; background-color:#999999;">
					<!-- 
						Aqui van los campos del formulario... 
					-->

					<%
						dim k, campo, con, sqlString, t, indicefondo, multiplicador, cedula
						
						sqlString = "SELECT Equipo, Nombre, Imagen, EnJuego " & _
										  "FROM mundial_Equipos " & _
										 "WHERE Equipo <> '-' " & _
									 "ORDER BY EnJuego DESC, NOMBRE ASC;"
						
						set con = Server.CreateObject("ADODB.Connection")
						con.open Application("Conn")				
						set t = con.execute(sqlString)
						
						if (t.bof or t.eof) then %>
							<tr>
								<td colspan="4" align="center">
									No encuentro equipos en el sistema!
								</td>
							</tr> <%
						else
							indicefondo = 1
							multiplicador = -1					
						%>
						
							<tr>
								<td colspan="4" align="center">
									<br />

									<table style="width: 90%; border-style: none; padding: 5px; border-spacing: 0;">
										<tr style="background-color:#666666;">
											<td colspan="2" style="font-size:18px; color:#FFFFFF; text-align:center;">Equipo</td>						
											<td style="font-size:18px; color:#FFFFFF; text-align:center;">Estado</td>
											<td style="font-size:18px; color:#FFFFFF; text-align:center;">Accion</td>
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
																						
											<td style="width: 10%; border-style: none; padding:0; border-spacing: 0; text-align: center;">
												<img src="<% 
													if t("EnJuego") = 1 then
														response.write "imagenes/banderas/"
													else
														response.write "imagenes/banderas2/"
													end if

													response.write t("Imagen")
												%>" style="border-style: none; width: 50px;">
											</td>
												
											<td style="text-align: left; width: 30%; border-style: none; padding: 15px; border-spacing: 0;">
												<%
													response.write "<span style='font-size:18px; color:"
													
													if t("EnJuego") = 1 then
														response.write "#000000;font-weight:bold;"
													else
														response.write "#B1B1B1;"
													end if
													
													response.write "'>" & t("Nombre") & "</span>"
												%>
											</td>
											
											<td style="text-align: center; width: 30%; border-style: none; padding: 15px; border-spacing: 0;">
												<%
													response.write "<span style='font-size:18px; color:"
													
													if t("EnJuego") = 1 then
														response.write "#000000;'>En Juego"
													else
														response.write "#B1B1B1;'>Eliminado"
													end if
													
													response.write  "</span>"
												%>
											</td>
											
											<td style="text-align: center; width: 30%; border-style: none; padding: 15px; border-spacing: 0;">
												<a href="<%
														response.write "actualizar_equipos.asp?e=" & t("Equipo") & "&s="
												
														if t("EnJuego") = 1 then
															response.write "0"
														else
															response.write "1"
														end if
												%>" <%
												if t("EnJuego") = 0 then
													response.write "class='eliminado'"
												end if
												%> style="font-size:18px;">
													<%
														if t("EnJuego") = 1 then
															response.write "Deshabilitar"
														else
															response.write "Habilitar"
														end if
													%>
												</a>
											</td>
										</tr>
									<%
											t.movenext
										loop until t.eof
									end if

									%>					
									</table>

									<br />
								</td>
							</tr>

						<%
							t.close: set t=nothing
							con.close: set con = nothing					
						%>
				</table>
		<%
			end if
		%>

		<script type="text/javascript">
            function Requery() {
                document.getElementById("formulario").submit();
            }
					
			function VerPolla(Codigo) {
				window.location = "ver.asp?t=" + Codigo
			}
		</script>		
		<!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>	
</html>