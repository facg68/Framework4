<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
		<title>Polla Mundial</title>
		
		<style type="text/css">
			body {
				background-color:#9F9F9F;
				background-image:url(Imagenes/fondo.jpg);
			}
			
			td { text-align:center; }
		
			.bandera {
				font-family:Verdana, Arial, Helvetica, sans-serif;
				font-size:10px;
				background-color:#000000;
				color:#FFFFFF;
			}
			
			a:link 		{ font-family:Arial, Helvetica, sans-serif; font-size:10px; color:#000; text-decoration: none; }
			a:visited 	{ font-family:Arial, Helvetica, sans-serif; font-size:10px; color:#000; text-decoration: none; }
			a:hover 	{ font-family:Arial, Helvetica, sans-serif; font-size:10px; color:#FF0000; text-decoration:underline; }
			
			.puntos {
				font-family:Verdana, Arial, Helvetica, sans-serif;
				font-size:14px;
				font-weight:bold;
				color:#FFF;
			}
			
			.fondo0 {
				background-color:#EEFCE4;
				font-family:Verdana, Arial, Helvetica, sans-serif;
				font-size:11px;
				color:#000;
			}
			
			.fondo1 {
				background-color:#F0EED0;
				font-family:Verdana, Arial, Helvetica, sans-serif;
				font-size:11px;
				color:#000;
			}
			
			.stretch{
			  background-size: 52px 34px;
			  background-repeat:no-repeat;
			}
			
			table {
				padding:0;
				border-spacing:0;
			}	

			::-webkit-scrollbar {
			    width: 8px;
			}

			::-webkit-scrollbar-button {
			    width: 8px;
			    height:5px;
			}

			::-webkit-scrollbar-track {
			    background:#eee;
			    border: thin solid lightgray;
			    box-shadow: 0px 0px 3px #dfdfdf inset;
			    border-radius:10px;
			}
			
			::-webkit-scrollbar-thumb {
			    background:#999;
				border: thin solid gray;
			    border-radius:10px;
			}

			::-webkit-scrollbar-thumb:hover {
			    background:#7d7d7d;
			} 		
			
			tr:not(:last-child) { border: none !important; }
		</style>
		
		<%
			Function NombrePais(CodEquipo)
				sqlString = "SELECT Nombre FROM dbo.mundial_Equipos WHERE Equipo = '" & CodEquipo & "';"
							
				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)
				
				if not (t.eof or t.bof) then
					NombrePais = t("Nombre")
				end if
			
				t.close: set t=nothing
				con.close: set con=nothing
			End Function
		%>
	</head>
	
	<body onload="Presentar()">
		<div style="display:none;">	
			<%
				Dim Tiquete, k, campo, con, sqlString, t, Pais
				Dim Nombre, Estatus, Total, multiplicador, indiceFondo 
				
				Pais = request.QueryString("e")
				Tiquete = Request.QueryString("t")
				
				sqlString = "Exec dbo.mundial_PollaPuntos '" & Tiquete & "';"
				
				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)
				
				if not(t.bof or t.eof) then
					Nombre = t("Nombre")
					Estatus = t("Estatus")
					Total = 0
					
					do
						%>
							<input id='CodEquipo<%= t("Posicion") %>'		value='<%= t("CodEquipo") %>' />							
							<input id='NombreEquipo<%= t("Posicion") %>'	value='<%= t("NombreEquipo") %>' />	
							<input id='Bandera<%= t("Posicion") %>'			value='<%= t("Imagen") %>' />
							<input id='Resultado<%= t("Posicion") %>'		value='<%= t("Resultado") %>' />
							<input id='Ganador<%= t("Posicion") %>'			value='<%= t("Ganador") %>' />
							<input id='Puntos<%= t("Posicion") %>'			value='<%= t("Puntaje") %>' />
							<input id='EnJuego<%= t("Posicion") %>'			value='<%= t("EnJuego") %>' />						
						<%
						
							Total = Total + t("Puntaje")
						t.movenext
					loop until (t.eof)
				end if
				
				t.close: set t =nothing
			%>		
		</div>

		<table width="95%" style="margin-left: auto; margin-right: auto; padding: 0px; border-spacing; 0px; background-color:#000000;">		
			<tr>
				<td><img src="Imagenes/header.jpg" style="border-style: none; width: 100%;" /></td>
			</tr>

			<tr>
				<td>
					<table width="100%" style="margin-left: auto; margin-right: auto; padding: 0px; border-spacing; 0px;">
						<tr>
							<td width="95%" style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 18px; text-align: center; color:rgb(255, 255, 255);">
								<strong>USUARIOS CON <%= UCase(NombrePais(Pais)) %> EN SUS POLLAS</strong>
								<%
									if Tiquete <> "" then
										response.write "<br />" & UCase(Nombre) & " / <span style='color:#FFFF00;'> CODIGO UNICO: " & Tiquete & "</span>"
										
										response.Write "&nbsp;&nbsp;&nbsp;[  Puntos: " & Total & "  ]  "
										
										if Estatus = 0 then
											response.write "&nbsp;&nbsp;&nbsp;-- sin activar --"
										end if
									else
										response.write "&nbsp;"
									end if
								%>								
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
					<table style="width: 100%; border-style: none;">
						<tr>
							<td width="20%" style="vertical-align:top;">
								<table style="width: 100%; border-style: none; padding: 0px: border-spacing: 0px;">
									<tr>
										<td style="background-color:#999999; color:#FFFFFF; font-family:Verdana, Arial, Helvetica, sans-serif;">
											<%= NombrePais(Pais)%>
										</td>
									</tr>
								
									<tr>
										<td>
											<div style="height:660px; overflow:auto;">
												<table style="width: 100%; border-style: none; padding: 0px: border-spacing: 0px;">
													<%
														sqlString = "Exec dbo.mundial_TablaPuntosPaisesIncluidos '" & Pais & "';"
														
														set con = server.CreateObject("ADODB.Connection")
														con.open Application("Conn")
														set t = con.execute(sqlString)
														
														if not(t.bof or t.eof) then
															indicefondo = 1
															multiplicador = -1
														
															do
																indicefondo = multiplicador + indiceFondo 
																multiplicador = -1 * multiplicador
																
																%>
																
																<tr class="fondo<%= indicefondo %>">
																	<td width="80%" style="text-align:left; padding: 10px; font-size: 12px;">
																		<span onclick="VerPolla('<%= t("Secuencia") %>', '<%= Pais %>')">
																			<%= Ucase(t("Nombre")) %>
																		</span>
																	</td>
																	
																	<td width="10%" style="text-align:center; padding: 10px; font-size: 12px;">
																		<a href="#" onclick="VerPolla('<%= t("Secuencia") %>', '<%= Pais %>')">
																			<%= t("PuntajeTotal") %>
																		</a>
																	</td>
																	
																	<td width="10%" style="vertical-align:center middle; padding: 10px;">
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
											</div>				
										</td>
									</tr>
								</table>
							</td>			
				
							<td width="80%" style="border-style: none; vertical-align:top;">
								<table style="margin-left: auto; margin-right: auto; width: 100%; background-image:url(Imagenes/estadio3_res.jpg); background-size: 100% 680px; border-style: none;">
									<tr>
										<td width="11%;">
										<td width="11%;">
										<td width="11%;">
										<td width="11%;">
										<td width="12%;">
										<td width="11%;">
										<td width="11%;">
										<td width="11%;">
										<td width="11%;">
									</tr>
								
									<tr>
										<td height="80">
											<label id="Fase1Nombre1" class="bandera"></label>							
											<table id="Fase1Bandera1" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase1Bandera1r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>
											<label id="Fase1Resultado1" class="puntos"></label>
										</td>
										
										<td rowspan="2">
											<label id="Fase2Nombre1"  class="bandera"></label>							
											<table id="Fase2Bandera1" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase2Bandera1r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>							
											<label id="Fase2Resultado1" class="puntos"></label>
										</td>
										
										<td>&nbsp;</td>
										<td>&nbsp;</td>
										
										<td rowspan="2">&nbsp;</td>
										<td>&nbsp;</td>
										<td>&nbsp;</td>
										
										<td rowspan="2">
											<label id="Fase2Nombre5" class="bandera"></label>								
											<table id="Fase2Bandera5" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase2Bandera5r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>							
											<label id="Fase2Resultado5" class="puntos"></label>
										</td>
										
										<td>
											<label id="Fase1Nombre9" class="bandera"></label>
											<table id="Fase1Bandera9" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase1Bandera9r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>
											<label id="Fase1Resultado9" class="puntos"></label>
										</td>
									</tr>
					
									<tr>
										<td height="80">
											<label id="Fase1Nombre2" class="bandera"></label>
											<table id="Fase1Bandera2" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase1Bandera2r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>
											<label id="Fase1Resultado2" class="puntos"></label>								
										</td>
										
										<td rowspan="2">
											<label id="Fase3Nombre1" class="bandera"></label>							
											<table id="Fase3Bandera1" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase3Bandera1r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>							
											<label id="Fase3Resultado1" class="puntos"></label>								
										</td>
										
										<td>&nbsp;</td>
										<td>&nbsp;</td>
										
										<td rowspan="2">
											<label id="Fase3Nombre3" class="bandera"></label>								
											<table id="Fase3Bandera3" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase3Bandera3r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>
											<label id="Fase3Resultado3" class="puntos"></label>									
										</td>
										
										<td>
											<label id="Fase1Nombre10" class="bandera"></label>
											<table id="Fase1Bandera10" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase1Bandera10r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>														
											<label id="Fase1Resultado10" class="puntos"></label>
										</td>
									</tr>
					
									<tr>
										<td height="80">
											<label id="Fase1Nombre3" class="bandera"></label>
											<table id="Fase1Bandera3" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase1Bandera3r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>
											<label id="Fase1Resultado3" class="puntos"></label>								
										</td>
										
										<td rowspan="2">
											<label id="Fase2Nombre2" class="bandera"></label>							
											<table id="Fase2Bandera2" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase2Bandera2r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>							
											<label id="Fase2Resultado2" class="puntos"></label>								
										</td>
										
										<td>&nbsp;</td>
										
										<td align="center">
											<label id="Fase5Nombre1" class="bandera"></label>							
											<table id="Fase5Bandera1" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase5Bandera1r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>
											<label id="Fase5Resultado1" class="puntos"></label><br />								
										</td>
										
										<td>&nbsp;</td>
										
										<td rowspan="2">
											<label id="Fase2Nombre6" class="bandera"></label>							
											<table id="Fase2Bandera6" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase2Bandera6r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>							
											<label id="Fase2Resultado6" class="puntos"></label>
										</td>
										
										<td>
											<label id="Fase1Nombre11" class="bandera"></label>
											<table id="Fase1Bandera11" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase1Bandera11r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>
											<label id="Fase1Resultado11" class="puntos"></label>
										</td>
									</tr>
					
									<tr>
										<td height="80">
											<label id="Fase1Nombre4" class="bandera"></label>							
											<table id="Fase1Bandera4" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase1Bandera4r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>
											<label id="Fase1Resultado4" class="puntos"></label>
										</td>
										
										<td>&nbsp;</td>
										
										<td rowspan="2">
											<label id="Fase4Nombre1" class="bandera"></label>							
											<table id="Fase4Bandera1" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase4Bandera1r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>							
											<label id="Fase4Resultado1" class="puntos"></label>
										</td>
										
										<td>&nbsp;</td>
										
										<td rowspan="2">
											<label id="Fase4Nombre2" class="bandera"></label>								
											<table id="Fase4Bandera2" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase4Bandera2r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>							
											<label id="Fase4Resultado2" class="puntos"></label>
										</td>
										
										<td>&nbsp;</td>
										
										<td>
											<label id="Fase1Nombre12" class="bandera"></label>								
											<table id="Fase1Bandera12" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase1Bandera12r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>
											<label id="Fase1Resultado12" class="puntos"></label>								
										</td>
									</tr>
					
									<tr>
										<td height="80">
											<label id="Fase1Nombre5" class="bandera"></label>							
											<table id="Fase1Bandera5" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase1Bandera5r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>
											<label id="Fase1Resultado5" class="puntos"></label>
										</td>
										
										<td rowspan="2">
											<label id="Fase2Nombre3" class="bandera"></label>							
											<table id="Fase2Bandera3" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase2Bandera3r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>							
											<label id="Fase2Resultado3" class="puntos"></label>
										</td>
										
										<td>&nbsp;</td>
										<td>&nbsp;</td>
										<td>&nbsp;</td>
										
										<td rowspan="2">
											<label id="Fase2Nombre7" class="bandera"></label>								
											<table id="Fase2Bandera7" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase2Bandera7r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>							
											<label id="Fase2Resultado7" class="puntos"></label>								
										</td>
										
										<td>
											<label id="Fase1Nombre13" class="bandera"></label>								
											<table id="Fase1Bandera13" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase1Bandera13r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>
											<label id="Fase1Resultado13" class="puntos"></label>
										</td>
									</tr>
					
									<tr>
										<td height="80">
											<label id="Fase1Nombre6" class="bandera"></label>							
											<table id="Fase1Bandera6" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase1Bandera6r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>							
											<label id="Fase1Resultado6" class="puntos"></label>
										</td>
										
										<td rowspan="2">
											<label id="Fase3Nombre2" class="bandera"></label>							
											<table id="Fase3Bandera2" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase3Bandera2r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>							
											<label id="Fase3Resultado2" class="puntos"></label>
										</td>
										
										<td colspan="3" rowspan="3">&nbsp;</td>
										
										<td rowspan="2">
											<label id="Fase3Nombre4" class="bandera"></label>
											<table id="Fase3Bandera4" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase3Bandera4r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>
											<label id="Fase3Resultado4" class="puntos"></label>								
										</td>
										
										<td>
											<label id="Fase1Nombre14" class="bandera"></label>								
											<table id="Fase1Bandera14" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase1Bandera14r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>								
											<label id="Fase1Resultado14" class="puntos"></label>
										</td>
									</tr>
					
									<tr>
										<td height="80">
											<label id="Fase1Nombre7" class="bandera"></label>							
											<table id="Fase1Bandera7" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase1Bandera7r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>							
											<label id="Fase1Resultado7" class="puntos"></label>
										</td>
										
										<td rowspan="2">
											<label id="Fase2Nombre4" class="bandera"></label>							
											<table id="Fase2Bandera4" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase2Bandera4r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>							
											<label id="Fase2Resultado4" class="puntos"></label>
										</td>
										
										<td rowspan="2">
											<label id="Fase2Nombre8" class="bandera"></label>								
											<table id="Fase2Bandera8" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase2Bandera8r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>							
											<label id="Fase2Resultado8" class="puntos"></label>
										</td>
										
										<td>
											<label id="Fase1Nombre15" class="bandera"></label>								
											<table id="Fase1Bandera15" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase1Bandera15r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>								
											<label id="Fase1Resultado15" class="puntos"></label>
										</td>
									</tr>
					
									<tr>
										<td height="80">
											<label id="Fase1Nombre8" class="bandera"></label>							
											<table id="Fase1Bandera8" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase1Bandera8r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>							
											<label id="Fase1Resultado8" class="puntos"></label>
										</td>
										
										<td>&nbsp;</td>
										<td>&nbsp;</td>
										
										<td>
											<label id="Fase1Nombre16" class="bandera"></label>								
											<table id="Fase1Bandera16" border="0" style="background-image:url(Banderas/bandera_def.png);" width="50px" height="34px" class="stretch" align="center">
												<tr><td>
													<img id="Fase1Bandera16r" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /> 
												</td></tr>
											</table>														
											<label id="Fase1Resultado16" class="puntos"></label>	
										</td>
									</tr>
								</table>			  
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		
		<%
			con.close: set con = nothing		
		%>

		<script type="text/javascript">
			function Presentar() {
				var limite = 32;
				var indice = 1;
				var bandera = "";
				var tabla;
				var imagen ="";
				
				for(fase=1;fase<=5;fase++) {
					limite/=2;
					
					for(equipo=1;equipo<=limite;equipo++) {
						bandera = "imagenes/banderas/";
						
						document.getElementById("Fase" + fase + "Nombre" + equipo).innerHTML = document.getElementById("NombreEquipo" + indice).value;
						tabla = document.getElementById("Fase" + fase + "Bandera" + equipo);
						
						//
						// En esta parte colocamos la bandera que el usuario puso en su apuesta,
						// ya sea activa (en juego) o desactivada (equipo ya no est� en el mundial)
						//
						
						if(document.getElementById("Resultado" + indice).value == '-' ) {
							if(document.getElementById("EnJuego" + indice).value != 1) bandera = "imagenes/banderas2/"; 
						}
						else {
							document.getElementById("Fase" + fase + "Resultado" + equipo).innerHTML = document.getElementById("Puntos" + indice).value;
						
							if(document.getElementById("Puntos" + indice).value == 0) { bandera = "imagenes/banderas2/"; }
							if(document.getElementById("EnJuego" + indice).value != 1) { bandera = "imagenes/banderas2/"; }
						}
						
						tabla.style.backgroundImage = "url(" + bandera + document.getElementById("Bandera" + indice).value + ")";
						
						//
						// Ahora vamos a verificar ganadores para colocarlos SOBRE
						// las banderas de la polla del jugador
						//
						
						if(document.getElementById("Resultado" + indice).value != '-' ) {
							if(document.getElementById("Resultado" + indice).value != document.getElementById("CodEquipo" + indice).value) {
								document.getElementById("Fase" + fase + "Bandera" + equipo + "r").src = "imagenes/banderas3/" + document.getElementById("Ganador" + indice).value.replace(".jpg", ".png");
							}
						}
					
						indice+=1;
					}
				}
			}
			
			function VerPolla(Secuencia, Pais) {
				window.location = "ver3.asp?t=" + Secuencia + "&e=" + Pais;
			}
		</script>		
	</body>
</html>