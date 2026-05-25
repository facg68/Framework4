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

        <style>
            body {
                background-color:#9F9F9F;
                background-image:url(Imagenes/fondo.jpg);
				overflow: auto;
            }
			
            td {
                font-family:Verdana, Arial, Helvetica, sans-serif;
            }
			
            #content h2 {
                text-align: left;
            }
			
			.fondo1 { background-color: rgb(255, 255, 255); }
			.fondo0 { background-color:#EBF8FE; }	
                       
            .res01 { background-color: rgb(255, 255, 255) }
            .res02 { background-color: rgb(231, 240, 216) }            
			
			a.mundial:link 		{ text-decoration: none; color:#000000; }
			a.mundial:visited 	{ text-decoration: none; color:#000000; }
			a.mundial:hover 	{ text-decoration: none; color:#000000; }
			a.mundial:active 	{ text-decoration: none; color:#000000; }	
			
			a.habilitado:link 			{ text-decoration: none; color:#000000; }
			a.habilitado:visited 		{ text-decoration: none; color:#000000; }
			a.habilitado:hover 			{ text-decoration: none; color:#000000; }
			a.habilitado:active 		{ text-decoration: none; color:#000000; }					
			
			a.deshabilitado:link 		{ text-decoration: none; color:#B1B1B1; }
			a.deshabilitado:visited 	{ text-decoration: none; color:#B1B1B1; }
			a.deshabilitado:hover 		{ text-decoration: none; color:#B1B1B1; }
			a.deshabilitado:active 		{ text-decoration: none; color:#B1B1B1; }	
			
			a.ganador:link 				{ text-decoration: none; color:#0000C4; }
			a.ganador:visited 			{ text-decoration: none; color:#0000C4; }
			a.ganador:hover 			{ text-decoration: none; color:#0000C4; }
			a.ganador:active 			{ text-decoration: none; color:#0000C4; }				
			
			.fondot0 {
				background-color:#EEFCE4;
				font-family:Verdana, Arial, Helvetica, sans-serif;
				font-size:9px;
				font-weight:bold;
				color:#000;
			}
			
			.fondot1 {
				background-color:#F0EED0;
				font-family:Verdana, Arial, Helvetica, sans-serif;
				font-size:9px;
				font-weight:bold;				
				color:#000;
			}	
			
			.topTen {
				background-color:#666666; 
				color:#FFFFFF; 
				font-family:Verdana, Arial, Helvetica, sans-serif; 
				font-size:12px;
				font-weight:bold;
				text-align:center;
			}
			
			.borderTableTopTen {
				border-style:solid;
				border-width:1px;
				border-color:#666666;
			}	
			
			.borderEstadisticas {
				border-style:solid;
				border-width:1px;
				border-color:#1F65BA;
			}	
			
			tr:not(:last-child) { border: none !important; }
        </style>
		
		<%
			Function MaximoPorcentaje()
				sqlString = "exec dbo.mundial_Porcentajes;"
							
				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)
				
				if not (t.eof or t.bof) then
					MaximoPorcentaje = int(t("Porcentaje"))
				else
					MaximoPorcentaje = 0
				end if
			
				t.close: set t=nothing
				con.close: set con=nothing
			end Function	
			
			function PollaActiva()
				sqlString = "SELECT Estatus FROM dbo.mundial_Estatus WHERE Codigo = 'Polla';"
							
				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)
				
				if not (t.eof or t.bof) then
					PollaActiva = t("Estatus")
				end if
			
				t.close: set t=nothing
				con.close: set con=nothing
			end function	

			function Finalizada()
				sqlString = "SELECT Estatus FROM dbo.mundial_Estatus WHERE Codigo = 'Finalizada';"
							
				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)
				
				if not (t.eof or t.bof) then
					Finalizada = t("Estatus")
				end if
			
				t.close: set t=nothing
				con.close: set con=nothing
			end function	

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
			
			function Acumulado()
				sqlString = "SELECT dbo.mundial_CuantoEnJuego() AS Cuanto;"
							
				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)
				
				if not (t.eof or t.bof) then
					Acumulado = t("Cuanto")
				end if
			
				t.close: set t=nothing
				con.close: set con=nothing
			end function	

			function valorBoleto()
				sqlString = "SELECT Estatus FROM mundial_Estatus WHERE Codigo = 'Boleto';"

				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)

				valorBoleto = FormatNumber(t("Estatus"))

				t.close: set t = nothing
				con.close: set con = nothing
			end function	

			function PollasActivadas()
				sqlString = "SELECT ISNULL(COUNT(*), 0) AS Cuantos FROM mundial_Apuestas_Enc WHERE (Estatus = 1);"

				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)

				PollasActivadas = t("Cuantos")

				t.close: set t = nothing
				con.close: set con = nothing
			end function	

			function HoraNumerica2Hora(Hora)
				dim horas, minutos, sufijo

				sufijo = "A.M."

				if len(Hora) = 4 then
					horas = left(hora, 2)
					minutos = right(hora, 2)

					if cInt(horas) > 12 then
					horas = right("00" & (horas - 12), 2)            
					sufijo = "P.M."
					end if

					HoraNumerica2Hora = cInt(horas) & ":" & minutos & " " & sufijo
				else
					HoraNumerica2Hora = ""
				end if
			end function			

			function FechaNumerica2Fecha(Fecha)
				dim dia, mes, amo

				if len(Fecha) = 8 then
					dia = right(fecha, 2)
					mes = mid(fecha, 5, 2)
					amo = left(fecha, 4)

					FechaNumerica2Fecha = dia & "/" & mes & "/" & amo
				else
					FechaNumerica2Fecha = ""
				end if			
			end function

			function FechaNumerica2FechaMin(Fecha)
				dim dia, mes

				if len(Fecha) = 8 then
					dia = right(fecha, 2)
					mes = mid(fecha, 5, 2)

					FechaNumerica2FechaMin = dia & "/" & mes 
				else
					FechaNumerica2FechaMin = ""
				end if			
			end function

			function DetalleJuegos(Equipo)
				dim con, t, sqlString, p1, p2
				dim label1, label2, linea

				sqlString = "exec mundial_Eliminatorias_JuegosPaises '" & Equipo & "'"

				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)

					if not (t.bof or t.eof) then
						DetalleJuegos = "&#10;"

						Do
							p1 = "0"
							p2 = "0"

							if t("Puntos1") > 0 then p1 = "+" & t("Puntos1")
							if t("Puntos2") > 0 then p2 = "+" & t("Puntos2")

							label1 = t("Equipo1") & ": " & t("Goles1") & " [" & p1 & "]"
							label2 = t("Equipo2") & ": " & t("Goles2") & " [" & p2 & "]"

							if t("CodEquipo1") = Equipo then
								linea = label1 & " - " & label2
							else
								linea = label2 & " - " & label1
							end if

							DetalleJuegos = DetalleJuegos & FechaNumerica2FechaMin(t("Fecha")) & " " & linea

							if t("Penales") = 1 then
								DetalleJuegos = DetalleJuegos & " (Penales)"
							end if

							DetalleJuegos = DetalleJuegos & "&#10;&#10;"

							t.MoveNext
						Loop Until t.eof
					end if

				t.close: set t = nothing
				con.close: set con = nothing
			end function			
			
			Sub Encabezado()
				%>
					<table width="100%" style="padding: 0; border-spacing: 0; background-color:#000000; color:#FFFFFF; border-style: none;">
						<tr>
							<td colspan="5">
								<img src="imagenes/header.jpg" style="width:100%; border-style: none;">
							</td>
						</tr>

						<tr>
							<td width="2%">&nbsp;</td>

							<td width="29%" style="text-align:left;font-size:20px; font-family:Verdana, Arial, Helvetica, sans-serif; padding: 10px;">
								<%
									if Finalizada() = 1 then
										Response.write "Juego Finalizado"
									else
										if PollaActiva() = 1 then	
											Response.write "Creacion de Pollas"
										else	
											Select Case Etapa()
												Case 1: Response.write "Fase de Grupos"
												Case 2: Response.write "Octavos de Final"
												Case 3: Response.write "Cuartos de Final"
												Case 4: Response.write "Semifinal"
												Case 5: Response.write "Final"
											End Select
										end if
									end if
								%>
							</td>
							
							<td width="38%" style="text-align:center;font-size:22px; font-family:Verdana, Arial, Helvetica, sans-serif; padding: 10px;">
								Total: $<%= Acumulado() %> (<%= PollasActivadas() %> Pollas) 
							</td>
							
							<td width="29%" style="font-size:14px; font-family:Verdana, Arial, Helvetica, sans-serif; text-align:right; padding: 10px;">
								<form name="Consultar" method="post" action="consultar.asp">
									<span  style="width: 30%;">Cedula: </span>
									<input class="field tiny" id="txtCedula" name="txtCedula" type="text" size="15" maxlength="15">
									<button type="submit" class="form-btn azul normal">Ver</button>
								</form>
							</td>

							<td width="2%">&nbsp;</td>
						</tr>
					</table>
				<%			
			End Sub

			Sub Participacion()
				%>
					<table width="100%" cellpadding="15" cellspacing="0">
						<tr style="background-color:#F8F8F8;">
							<td style="padding: 15px;">				
								<h2>Reglas para Participar</h2>
								
								<br>
								
								<div style="font-family:Verdana, Arial, Helvetica, sans-serif; font-size:13px; text-align:left;">
									<strong>Gracias por participar en nuestra Polla 2022!</strong>
									
									<br><br>
									
									En la siguiente pagina debe llenar el cuadro de futbol con los ganadores.<br><br>
									
									Al terminar, se le pedira que escriba su nombre, numero de cedula, departamento al que pertenece y el numero de la extension 
									donde podamos localizarlo en caso que sea uno de los ganadores!<br><br>
									
									Al llenar sus datos, el sistema le dara un <strong>Codigo Unico</strong> que identificara su apuesta. Debe apuntar este codigo ya que, sin 
									el mismo, no podra participar en el concurso.<br><br>
									
									Usted puede crear todas las "pollas" que quiera, y cada una de ellas le otorgara un codigo distinto. Lleve estos codigos a la <strong>Sede
									del Juego</strong> para "activar" sus codigos. Si su codigo no se activa, su "polla" NO sera eliminada del sistema, pero no competira 
									por el premio final.<br><br>
									
									Cada activacion <strong>tiene un valor de <%= valorBoleto() %> balboas</strong> y, al ser activado cada codigo, se le entregara una copia <strong>firmada</strong> en papel 
									de su polla y su codigo activado. Debe conservar este "boleto" para reclamar los premios al final del torneo.<br><br>

									Un 10% de cada boleto sera destinado al pago y mantenimiento del servidor donde esta aplicaci&oacute;n ha sido instalada. El resto ser&aacute; acumulado como premio para los ganadores.<br/><br/>
									
									Una copia impresa de cada polla se guardara en nuestras oficinas para garantizar la transparencia del juego.<br><br>
									
									<b>LAS POLLAS PUEDEN SER CREADAS HASTA EL 11 DE NOVIEMBRE DEL 2022 Y EL PERIODO DE ACTIVACION FINALIZA EL 18 DE NOVIEMBRE DEL 2022.</b><br><br>
									
									<span style="color:#CC0000; font-weight:bold;">
										NOTA: Queremos que todos los usuarios participen de nuestro juego, por lo tanto, ninguna de las Pollas sera eliminada del 
										Sistema, asi que puedes crear tu(s) polla(s) sin complicaciones. Solo recuerda que si la polla no esta activa, no participara de los premios!
										<br><br>
										GRACIAS POR EL APOYAR NUESTRAS INICIATIVAS!!!
										<br><br>
									</span>
								</div>	
							</td>
						</tr>
					</table>		
				<%
			End Sub
			
			Sub ReglaPuntajes() 
				%>
					<table width="100%" cellpadding="15" cellspacing="0">
						<tr style="background-color:#F8F8F8;">
							<td style="padding: 15px;">				
								<br>
								
								<h3>Como se otorgan los puntos:</h3>
								
								<div style="font-family:Verdana, Arial, Helvetica, sans-serif; font-size:13px; text-align:left;">
									<ul>
										<li>Cada Fase del Torneo le adjudica puntos.</li>
										<br>
										<li>Cada Equipo que acierte en los <strong>Octavos de Final</strong> (en la posicion correcta) le otorga <strong>2 puntos</strong>.</li>
										<br>
										<li>Cada Equipo que acierte en los <strong>Octavos de Final</strong> (en la posicion incorrecta) le otorga <strong>1 punto</strong>.</li>
										<br>
										<li>Cada Equipo que acierte en los <strong>Cuartos de Final </strong>(en la posicion correcta) le otorga <strong>4 puntos</strong>.</li>
										<br>
										<li>Cada Equipo que acierte en los <strong>Cuartos de </strong><strong>Final</strong> (en la posicion incorrecta) le otorga <strong>2 puntos</strong>.</li>
										<br>
										<li>Cada Equipo que acierte en la <strong>Semi-Final </strong>(en la posicion correcta) le otorga <strong>6 puntos</strong>.</li>
										<br>
										<li>Cada Equipo que acierte en la <strong>Semi-Final</strong> (en la posicion incorrecta) le otorga <strong>3 puntos</strong>.</li>
										<br>
										<li>Cada Equipo que acierte en la <strong>Final</strong> (en la posicion correcta) le otorga <strong>8 puntos</strong>.</li>
										<br>
										<li>Cada Equipo que acierte en la <strong>Final</strong> (en la posicion incorrecta) le otorga <strong>4 puntos</strong>.</li>
										<br>									
										<li>Acertar el Equipo que <strong>Gane la Copa</strong> le otorga <strong>10 puntos</strong>!</li>
									</ul>						
								</div>
								
								<br>

								<h3>Como se Paga al Ganador (o gandores):</h3>
								
								<div style="font-family:Verdana, Arial, Helvetica, sans-serif; font-size:13px; text-align:left;">
									<ul>
										<li><strong>Si s&oacute;lo una persona predice correctamente al ganador del torneo, se lleva el acumulado</strong> (sin tomar en cuenta los puntos) </li>
										<br>
										<li>Si <strong>m&aacute;s de una persona</strong> predice correctamente al ganador del torneo, <strong>se divide el total acumulado</strong> de la siguiente forma:</li>
										<br>
										<ul>
											<li>La persona que tenga <strong>el mayor puntaje</strong> se lleva el acumulado.</li>
											<br>
											<li>Si <strong>varias personas</strong> tienen un &quot;mayor puntaje igual&quot;, <strong>se divide el acumulado entre esas personas</strong>.</li>
											<br>
										</ul>															  
										<br>

										<li>Si <strong>ninguna persona</strong> predice correctamente al ganador, <strong>se toman en cuenta los puntajes</strong>, de la siguiente forma: 
											<br><br>
											
											<ul>
												<li>Si <strong>una s&oacute;la persona</strong> acumula m&aacute;s puntos, &eacute;sta se lleva el acumulado  </li>
												<br>
												<li>Si <strong>varias personas</strong> tienen un &quot;mayor puntaje igual&quot;, <strong>se divide el acumulado entre esas personas</strong>.</li>
												<br>
											</ul>
										</li>

										<br>
									
										<li><strong>Al final del torneo se publicar&aacute; la lista de ganadores en &eacute;sta P&aacute;gina</strong>.</li>
										<br>
									</ul>	
								</div>			
							</td>
						</tr>
					</table>
				<%
			End Sub

			Sub PollaClick()
				%>
					<table width="100%" cellpadding="15" cellspacing="0">
						<tr style="background-color:#E2E2E2;">
							<td style="padding: 15px; text-align: center;">
								<a class="mundial" href="seleccion.asp" style="font-size:16px; font-weight:bold;">Haga Click Aqui Para Empezar!</a>
							</td>
						</tr>
					</table>				
				<%
			End Sub

			Sub ListaDeGanadores()
				sqlString = "SELECT Nombre, Polla, Premio FROM mundial_ListaDeGanadores;"
							
				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)
				
				if not (t.eof or t.bof) then
				%>
		
					<br><br>

					<div style="margin: 0 auto; width: 60%; background-color:#F8F8F8;">
						<table width="100%" style="text-align: center" cellpadding="5" cellspacing="0" class="borderEstadisticas">
							<tr>
								<td colspan="3"
									style="padding: 6px; text-align: center; font-size: 20px; font-weight:bold; background-color: black; color: white;">
									Lista de Ganadores
								</td>
							</tr>
							
							<tr class="topTen">
								<td style="padding: 6px;">Ganador</td>
								<td style="padding: 6px;">Polla</td>
								<td style="padding: 6px;">Premio</td>
							</tr>
							
							<% Do %>
								<tr style="background-color: rgb(226, 235, 213);">
									<td style="padding: 6px;text-align: center">
										<a class="mundial" href="#" onClick="VerPolla('<%= t("Polla") %>')" class="ganador">
											<%= Ucase(t("Nombre")) %>
										</a>											
									</td>
									
									<td style="padding: 6px;text-align: center">
										<a class="mundial" href="#" onClick="VerPolla('<%= t("Polla") %>')" class="ganador">
											<%= Ucase(t("Polla")) %>
										</a>											
									</td>
									
									<td style="padding: 6px;text-align: center">
										<a class="mundial" href="#" onClick="VerPolla('<%= t("Polla") %>')" class="ganador">
											$<%= Ucase(t("Premio")) %>
										</a>											
									</td>
								</tr>
							<%
								t.movenext
							loop until (t.eof)		
							%>
							
						</table>
					</div>
				<%
									
					t.close: set t=nothing
					con.close: set con = nothing											
				end if			
			End Sub

			Sub MensajeNotificacionFinDeApuestas()
				%>
					<table style="width: 95%; padding: 0px; border-spacing: 0; margin-left: auto; margin-right: auto;">
						<tr style="background-color:#F8F8F8;">
							<td style="text-align: center;">
								<span style="font-size:12px;">							
									<br />
									El Periodo de creacion de las Pollas ha terminado y, de ahora en adelante, daremos seguimiento a los juegos mediante su codigo unico!<br><br>
									
									Para garantizar la transparencia de nuestro juego, todas las pollas son publicas, de modo que, a medida que se vayan acumulando puntos,<br />
                                    usted podra consultar las pollas que tengan mas puntaje para verificar sus equipos.
								</span>
							</td>
						</tr>			
					</table>

					<br />
				<%
			End Sub

			Sub EstadisticasParticipantes()
				Dim con, t, sqlString, indice, base, sw, swVerTopTen

				base = MaximoPorcentaje()
				sw = 0
				sqlString = "exec dbo.mundial_Porcentajes;"
				swVerTopTen = Request.QueryString("tt")

				if swVerTopTen = "" then
					if Etapa() > 2 then
						swVerTopTen = 1
					else
						swVerTopTen = 0
					end if
				end if

				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")				

				%>
					<br >

					<table style="width: 95%; padding: 0px; border-spacing: 0; margin-left: auto; margin-right: auto;">
						<tr style="background-color:#F8F8F8;">
							<td style="width: 73%; text-align: center;">
								<span style="font-size:14px;font-weight:bold; color: red;">
									Estos son los Equipos Seleccionados por los Apostadores. Consulte las Estadisticas libremente.
								</span>
								
								<br><br>

								<%
									set t = con.execute(sqlString)
									
									if not (t.eof or t.bof) then
										indice = 1
										%>
											<table width="100%" border="0" cellpadding="5" cellspacing="0" class="borderEstadisticas">
										<%
											do
										%>
												<tr class="fondo<%
													response.write sw
													if sw = 0 then sw = 1 else sw = 0
												%>" onClick="ver('<%= t("CodEquipo") %>')">
												
													<td width="25%" style="padding: 8px; text-align: right">
														<%
															response.write "<span class='"
																if t("EnJuego") = 1 then 
																	response.write "habilitado"  
																else
																	response.write "deshabilitado"
																end if
															response.write "'>"
															
															response.write "<span style='font-size:14px;'>"
															response.write t("Nombre")
															response.write "</span>"
															
															response.write "</span>"
															
															response.write "&nbsp;&nbsp;"
															
															response.write "<span href='#'>"
																response.write "<img src='"
																
																if t("EnJuego") = 1 then 
																	response.write "imagenes/banderas"
																else
																	response.write "imagenes/banderas2"
																end if
																
																response.write "/" & t("Imagen") & "' border='0' width='33' height='22'>"
															response.write "</span>"
															
															response.write "&nbsp;&nbsp;"														
														%>
													</td>
													
													<td width="75%" style="padding: 8px; text-align: left">
														<%
															response.write "<span href='#'>"
																response.write "<img src='Imagenes/barras/"
															
																if t("EnJuego") = 1 then
																	response.write "azul"
																else
																	response.write "gris"
																end if
																
																response.write ".jpg' width='" &  ((t("Porcentaje") * 450) / base) & "' height='30' border=0>"
															response.write "</span>"
															
															response.write "&nbsp;&nbsp;"
															
															response.write "<span style='font-size:12px;'>"
																response.write  t("Cuantos") & " (" & t("FullPorcentaje") & "%)"
															response.write "</span>"
														%>
														</span>
													</td>
												</tr>														
										<%
												'indice = indice + 1
												if indice > 7 then indice = 1
												
												t.movenext
											loop until (t.eof)		
											
											t.close: set t = nothing									
										%>
											</table>
										<%
									else
										response.write "Aun No Hay Estadisticas Disponibles!"
									end if
								%>
							</td>

							<td style="width: 2%;">
								&nbsp;
							</td>

							<td style="width: 25%;">
								<table style="width: 100%; border-spacing: 0px;" class="borderTableTopTen">
									<tr>
										<td colspan="4" style="border-style: none; padding: 7px; text-align: center;" class="topTen"
											onclick="CambiarTopTen(<% if swVerTopTen = 0 then response.write "1" else response.write "0" %>)">
											TOP 10 <%
												if swVerTopTen = 1 then
													response.write "EN JUEGO  -  (ver todo)"
												else
													response.write "  -  (aplicar filtro)"
												end if												
											%>
										</td>
									</tr>
								
									<%
										if swVerTopTen = 1 then
											sqlString = "Exec dbo.mundial_TopTen_P;"
										else
											sqlString = "Exec dbo.mundial_TopTen;"
										end if
										
										set t = con.execute(sqlString)
										
										if not(t.bof or t.eof) then
											indicefondo = 1
											multiplicador = -1
										
											do
												indicefondo = multiplicador + indiceFondo 
												multiplicador = -1 * multiplicador
												
												%>
												
												<tr class="fondot<%= indicefondo %>" onClick="VerPolla('<%= t("Secuencia") %>')">
													<td width="10%" style="border-style: none; padding: 7px; font-family: Helvetica; font-size: 12px; font-weight: normal;">
														<%
															response.write "<img src='"

															if t("EnJuego") = 1 then 
																response.write "imagenes/banderas"
															else
																response.write "imagenes/banderas2"
															end if

															response.write "/" & t("Bandera") & "' border='0' height='16px'>"
														%>
													</td>

													<td width="75%" style="border-style: none; padding: 7px; font-family: Helvetica; font-size: 12px; font-weight: normal; color: <%
														if t("EnJuego") = 1 then 
															response.write "rgb(0, 0, 0)"															
														else
															response.write "rgb(140, 140, 140)"
														end if						
													%>;">

														&nbsp;<%= Ucase(t("Nombre")) %>
													</td>
													
													<td width="10%" style="border-style: none; padding: 7px; font-family: Helvetica; font-size: 12px; font-weight: normal;color: <%
														if t("EnJuego") = 1 then 
															response.write "rgb(0, 0, 0)"															
														else
															response.write "rgb(140, 140, 140)"
														end if						
													%>;"">
														<%= t("PuntajeTotal") %>
													</td>
													
													<td width="5%" style="border-style: none; padding: 7px; vertical-align:middle;">
														<%= "<img src='Imagenes/estatus" & t("Estatus") & ".png' border='0' />" %>
													</td>
												</tr>
												
												<%
												t.movenext
											loop until t.eof
										end if

										t.close: set t=nothing										
									%>
								</table>							
							</td>
						</tr>
					</table>

					<br />
				<%

				con.close: set con = nothing				
			End Sub

			Sub HistorialPartidos(Etapa)
				dim con, t, sqlString, sw

				If Etapa = 6 then
					sqlString = "SELECT * FROM mundial_Historial " & _
							"ORDER BY Etapa DESC, Fecha DESC, Hora Desc;"				
				else
					sqlString = "SELECT * FROM mundial_Historial " & _
								"WHERE (Etapa = " & Etapa & ") " & _
							"ORDER BY Fecha DESC, Hora Desc;"
				end if

                sw = -1
							
				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)

				CuantosPartidos = 0
				
				if not (t.eof or t.bof) then
					response.write "<div style='width: 95%; margin: auto;'>"
                        %>
                            <table style="width:100%; border-spacing: 0px; border: solid 1px black; margin-left: auto; margin-right: auto;" class="res01">
                                <tr style="color: white; background-color: black; border-spacing: 0px;">
                                    <td colspan="9" style="text-align: center; padding: 10px;">
										Historial de Partidos
										<%
											Select Case Etapa
												Case 1: Response.write " - Fase de Grupos"
												Case 2: Response.write " - Octavos de Final"
												Case 3: Response.write " - Cuartos de Final"
												Case 4: Response.write " - Semifinal"
												Case 5: Response.write " - Final"
											End Select										
										%>
									</td>
                                </tr>               

                                <tr style="color: white; background-color: rgb(74, 74, 74); 
										   padding: 10px; border-spacing: 0px; font-size: 14px;
										   font-family: arial;">
                                    <td style="width: 15%; text-align: center; padding: 5px;">Etapa</td>
                                    <td style="width:  6%; text-align: center; padding: 5px;">Grupo</td>

                                    <td style="width: 15%; text-align: center; padding: 5px;">Fecha</td>
                                    <td style="width: 20%; text-align: left; padding: 5px;">Equipo 1</td>
                                    <td style="width:  6%; text-align: center; padding: 5px;">Goles</td>
                                    <td style="width:  6%; text-align: center; padding: 5px;">&nbsp;</td>
                                    <td style="width: 20%; text-align: left; padding: 5px;">Equipo 2</td>
                                    <td style="width:  6%; text-align: center; padding: 5px;">Goles</td>
                                    <td style="width:  6%; text-align: center; padding: 5px;">&nbsp;</td>
                                </tr>
                        <%
							Do
                                sw = -1 * sw
								CuantosPartidos = CuantosPartidos + 1

								%>
									<tr style="padding: 10px; border-spacing: 0px;" class="res0<%
                                        if sw > 0 then
                                            response.write "1"
                                        else
                                            response.write "2"
                                        end if
                                    %>">
										<td style="text-align: center; padding: 10px;">
											<%
												Select Case t("Etapa")
													Case 1
														Response.write "Fase de Grupos"
													Case 2
														Response.write "Octavos"
													Case 3
														Response.write "Cuartos"
													Case 4
														Response.write "Semi Final"
													Case 5
														Response.write "Final"
												End Select
											%>
										</td>

										<td style="text-align: center; padding: 10px;"><%= t("Grupo") %></td>
										
										<td style="font-size: 12px; text-align: center; padding: 10px;"><%= FechaNumerica2Fecha(t("Fecha")) & "<br/>" & HoraNumerica2Hora(t("Hora")) %></td>

										<td onClick="ver3('<%= t("Equipo1") %>')">
											<%
												bandera_1 = Replace(t("Bandera1"), "Banderas", "imagenes/banderas")

												if t("Goles1") <> t("Goles2") then
													if t("Goles1") > t("Goles2") then
														%>
															<img src="<%= bandera_1 %>" style="width: 33px; height: 22px; border-style: none;">
															&nbsp;&nbsp;
															<span style="color: rgb(3, 50, 168);"><%= t("Nombre1") %></span>
														<%
													else
														%>
															<img src="<%= Replace(bandera_1, "imagenes/banderas", "imagenes/banderas2") %>" style="width:33px; height: 22px; border-style: none;">
															&nbsp;&nbsp;
															<span style="color: rgb(150, 150, 150);"><%= t("Nombre1") %></span>
														<%
													end if
												else
													%>
														<img src="<%= bandera_1 %>" style="width: 33px; height: 22px; border-style: none;">
														&nbsp;&nbsp;
														<span style="color: rgb(3, 50, 168);"><%= t("Nombre1") %></span>
													<%
												end if
											%>
                                        </td>

										<td style="text-align: center; padding: 10px;" onClick="ver3('<%= t("Equipo1") %>')">
											<% 
												if t("Goles1") <> t("Goles2") then
													if t("Goles1") > t("Goles2") then
														%><span style="color: rgb(3, 50, 168);"><%= t("Goles1") %></span><%
													else
														%><span style="color: rgb(150, 150, 150);"><%= t("Goles1") %></span><%
													end if
												else
													%><span style="color: rgb(3, 50, 168);"><%= t("Goles1") %></span><%
												end if											
											%>
										</td>

										<td>&nbsp;</td>

										<td onClick="ver3('<%= t("Equipo2") %>')">
											<%
												bandera_2 = Replace(t("Bandera2"), "Banderas", "imagenes/banderas")

												if t("Goles1") <> t("Goles2") then
													if t("Goles2") > t("Goles1") then
														%>
															<img src="<%= bandera_2 %>" style="width: 33px; height: 22px; border-style: none;">
															&nbsp;&nbsp;
															<span style="color: rgb(3, 50, 168);"><%= t("Nombre2") %></span>
														<%
													else
														%>
															<img src="<%= Replace(bandera_2, "imagenes/banderas/", "imagenes/banderas2/") %>" style="width: 33px; height: 22px; border-style: none;">
															&nbsp;&nbsp;
															<span style="color: rgb(150, 150, 150);"><%= t("Nombre2") %></span>
														<%
													end if
												else
													%>
														<img src="<%= bandera_2 %>" style="width: 33px; height: 22px; border-style: none;">
														&nbsp;&nbsp;
														<span style="color: rgb(3, 50, 168);"><%= t("Nombre2") %></span>
													<%
												end if
											%>
                                        </td>

										<td style="text-align: center; padding: 10px;" onClick="ver3('<%= t("Equipo2") %>')">
											<%
												if t("Goles1") <> t("Goles2") then
													if t("Goles2") > t("Goles1") then
														%><span style="color: rgb(3, 50, 168);"><%= t("Goles2") %></span><%
													else
														%><span style="color: rgb(150, 150, 150);"><%= t("Goles2") %></span><%
													end if
												else
													%><span style="color: rgb(3, 50, 168);"><%= t("Goles2") %></span><%
												end if
											%>
										</td>

										<td style="text-align: center; padding: 10px;">
											<%
												if t("Penales") = 1 then
													response.write "<span style='color:rgb(161, 2, 2);'>P</span>"
												else
													response.write "&nbsp;"
												end if
											%>
										</td>										
									</tr>
								<%

								t.MoveNext
							Loop Until t.eof

						%>
							<tr style="color: white; background-color: rgb(0, 0, 0); padding: 10px; border-spacing: 0px;">
								<td colspan="9" style="width: 100%; text-align: center; padding: 5px;"><%= CuantosPartidos %> Partidos</td>
							</tr>
						<%

						response.write "</table>"
					response.write "</div>"
				end if
			
				t.close: set t=nothing
				con.close: set con=nothing
			End Sub

			Sub DibujarEstadoGrupo(Grupo)
				dim con, t, sqlString, sw

				sqlString = "exec mundial_eliminatorias_grupos '" & Grupo & "';"

				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)

				if (not t.bof or t.eof) then
					sw = -1			
					%>
						<table style="width:100%; border-spacing: 0px; border: solid 1px black; font-size: 14px;" class="res01">
							<tr>
								<td colspan="2" style="color: white; background-color: rgb(74, 74, 74); padding: 10px; border-spacing: 0px;">
									Grupo <%= Grupo %>
								</td>
							</tr>
							</tr>
							<%	
								Do
									sw = -1 * sw
							%>
								<tr class="res0<% 
									if sw > 0 then 
										response.write "1"
									else	
										response.write "2"
									end if
								%>" title="<%= DetalleJuegos(t("Equipo")) %>">
									<td style="width: 80%; Text-align: left; padding: 10px;">
										<%
											if t("EnJuego") = 1 then
												%>
													<img src="imagenes/banderas/<%= t("Imagen") %>" style="width: 33px; height: 22px; border-style: none;" onClick="ver3('<%= t("Equipo") %>')">
													&nbsp;&nbsp;	
													<span style="color: rgb(3, 50, 168); font-family: Arial Narrow; font-size: 18px;" onClick="ver3('<%= t("Equipo") %>')"><%= t("Nombre") %></span>
												<%
											else
												%>
													<img src="imagenes/banderas2/<%= t("Imagen") %>" style="width: 33px; height: 22px; border-style: none;" onClick="ver3('<%= t("Equipo") %>')">
													&nbsp;&nbsp;	
													<span style="color: rgb(150, 150, 150); font-family: Arial Narrow; font-size: 18px;" onClick="ver3('<%= t("Equipo") %>')"><%= t("Nombre") %></span>												
												<%
											end if
										%>
									</td>

									<td style="width: 20%; Text-align: center; padding: 10px;">
										<%
											if t("EnJuego") = 1 then
												response.write "<span style='color: rgb(3, 50, 168);'>" & t("Puntaje") & "</span>"
											else
												response.write "<span style='color: rgb(150, 150, 150);'>" & t("Puntaje") & "</span>"
											end if									
										%>
									</td>
								</tr>

							<%
									t.MoveNext
								Loop Until t.eof
							%>
						</table>
					<%
				else
					response.write "&nbsp;"
				end if

				t.close: set t = nothing
				con.close: set con=nothing
			End Sub

			Sub Grupos_Estatus()
				%>
					<table style="width:95%; border-spacing: 0px; border-style: none; margin-left: auto; margin-right: auto;">
						<tr style = "border-style: none; padding:0; border-spacing: 0;">
							<td style="width: 24%;">
								<% DibujarEstadoGrupo("A") %>
							</td>

							<td style = "width: 2%;">&nbsp;</td>

							<td style="width: 23%;">
								<% DibujarEstadoGrupo("B") %>
							</td>

							<td style = "width: 2%;">&nbsp;</td>

							<td style="width: 23%;">
								<% DibujarEstadoGrupo("C") %>
							</td>

							<td style = "width: 2%;">&nbsp;</td>

							<td style="width: 24%;">
								<% DibujarEstadoGrupo("D") %>
							</td>
						</tr>

						<tr>
							<td colspan="7">
								&nbsp;
							</td>
						</tr>

						<tr style = "border: none; padding:0; border-spacing: 0;">
							<td style="width: 24%;">
								<% DibujarEstadoGrupo("E") %>
							</td>

							<td style = "width: 2%;">&nbsp;</td>

							<td style="width: 23%;">
								<% DibujarEstadoGrupo("F") %>
							</td>

							<td style = "width: 2%;">&nbsp;</td>

							<td style="width: 23%;">
								<% DibujarEstadoGrupo("G") %>
							</td>

							<td style = "width: 2%;">&nbsp;</td>

							<td style="width: 24%;">
								<% DibujarEstadoGrupo("H") %>
							</td>
						</tr>
					</table>

					<br />						
				<%
			End Sub

			Sub Gracias()
				%>
					<table width="100%" cellpadding="15" cellspacing="0">
						<tr style="background-color:#F8F8F8;">
							<td style="padding: 15px; text-align: center;">
							<!--
								<span style="font-family:Arial, Helvetica, sans-serif; font-size:12px; color:#800000; font-weight:bold;">
									NOTA ACLARATORIA: NINGUN MIEMBRO DE NUESTRO DEPARTAMENTO PUEDE JUGAR POR EL ACUMULADO!
								</span>
								
								<br><br>
							-->
								<strong>Gracias por participar en nuestra Polla 2022 y Gracias por apoyarnos!</strong>

							</td>
						</tr>
					</table>
				<%
			End Sub
		%>
    </head>
	
    <body>
		<!-- #include virtual = "/core/includes/kernel/body.inc" -->

		<%
			Response.Cookies("mundial_key") = "menu_pollas"
		%>

		<table style="margin-left: auto; margin-right: auto;" width="98%" cellpadding="15" cellspacing="0">
			<tr style="background-color:#F8F8F8; vertical-align: top; padding:0; cell-spacing: 0;">
				<td style="padding: 15px; vertical-align: top; padding:0; cell-spacing: 0;">		
					<%
						Encabezado

						if Finalizada() = 1 then
							ListaDeGanadores
						end if

						if PollaActiva() = 1 then	
							Participacion
							ReglaPuntajes
							PollaClick
						else
							if Etapa() = 1 then	MensajeNotificacionFinDeApuestas

							EstadisticasParticipantes
							
							if Etapa() = 1 then Grupos_Estatus
							if Etapa() > 0 then HistorialPartidos(Etapa())
						end if
        
						Gracias
					%>					
					<br />
				</td>
			</tr>
		</table>    

		<br />
		
		<script type="text/javascript">
			function ver(CodigoEquipo) {
				window.location = "ver2.asp?e=" + CodigoEquipo
			}

			function ver3(CodigoEquipo) {
				window.location = "ver3.asp?e=" + CodigoEquipo
			}

			function VerPolla(Secuencia) {
				window.location = "ver.asp?t=" + Secuencia;
			}	

			function CambiarTopTen(sw) {
				window.location = "mundial.asp?tt=" + sw;
			}		
		</script>
		<!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>