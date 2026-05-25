<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->
        <%
            thisSystem = "mundial"
            thisProcess = "mundial.020"
            SysLockOut
        %>    

		<title>Actualizar Master</title>
		
		<style type="text/css">
			body {
				background-color: #9F9F9F;
				background-image:url(imagenes/fondo.jpg);
				overflow: auto;
			}
		
			td {
				font-family:Verdana, Arial, Helvetica, sans-serif;
				font-size:10px;
			}

			.centro { text-align:center; }
			
			a.menu:link, a.menu:visited { font-family:Verdana, Arial, Helvetica, sans-serif; font-size:16px; color:#FFF; text-decoration: none; }
			a.menu:hover, a.active 	{ font-family:Verdana, Arial, Helvetica, sans-serif; font-size:16px; color:#FFFF4F !important; font-weight:bold !important; text-decoration: none;}

			#Mapa td img:hover { cursor: pointer; }
			div#CopaCampeon img:hover, td img#Fase5Bandera1:hover { cursor: default; }

			tr:not(:last-child) { border: none !important; }
		</style>

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
		
			Function SePuedeActualizarMaster()
				dim con, t, sqlString
				
				sqlString = "SELECT Estatus FROM dbo.mundial_Estatus WHERE Codigo = 'Finalizada';"
				SePuedeActualizarMaster = 0
				
				set con = Server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)
				
				if t("Estatus") = 0 then
					SePuedeActualizarMaster = 1
				end if

				t.close: set t=nothing
				con.close: set con=nothing			
			End Function
		%>
	</head>

    <body onload="PresentarMaster(); CambiarGrupo('A'); toggleGrupos();">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

		<%
			dim con, sqlString, t, categoria, sw, secuencia, apendice, Grupo
			dim Codigo, Nombre, Posicion, indice, gr, key
		%>	
		
		<table style="width: 95%; margin-left: auto; margin-right: auto; padding: 0px; border-spacing: 0px; background-color: rgb(0, 0, 0);">	
			<tr>
				<td>
					<table style="width: 100%; padding: 0px; border-spacing: 0px; background-color: rgb(0, 0, 0);">
						<tr colspan ="3" style="background-color: rgb(0, 0, 0); padding: 0px; border-spacing: 0px;">
							<img src="Imagenes/header.jpg" style="width:100%; border-style: none;">
						</tr>

						<tr style="background-color: rgb(0, 0, 0); padding: 0px; border-spacing: 0px;">
							<td class ="centro" width="25%" onclick="toggleGrupos()" style="color:#FFFFFF; font-size:16px; padding: 15px;">
								<label id="toggleOcultarGrupos">VER GRUPOS</label>
							</td>
							
							<td class ="centro" width="50%" cstyle="color:#FFFFFF; font-size:16px; padding: 15px;">
								<div id="SeleccionGrupo" align="center" style="display:block;">
								GRUPOS: 
								<%
									gr = "ABCDEFGH"
									
									for secuencia = 1 to 8
										response.write "<a href='#' onClick=""" & "CambiarGrupo('" & mid(gr,secuencia,1) & "')" & """ id='grupo" & mid(gr,secuencia,1) & "' class='menu'>&nbsp;" & mid(gr,secuencia,1) & "&nbsp;</a>&nbsp;&nbsp;"
									next
								%>
								</div>
								
								<div id="ActualizarEquipos" align="center" style="display:none;">
									<a href="equipos.asp" class="menu">ACTUALIZAR EQUIPOS</a>
								</div>
							</td>
											
							<td class ="centro" width="25%" style="font-size:16px; padding: 15px;">
								<input style="background-color: black; color: white;" id="btnGrabar" type="button" value="Actualizar Master" onclick="PasarSubmit();" />				
							</td>					
						</tr>
					</table>

					<div id="Preliminares" align="center" style="display:block;">	

						<table width="100%" cellpadding="15" cellspacing="0">
							<tr style="background-color:#E2E2E2;">
								<td class ="centro" width="15%">
									<span style="font-size:10px">
										Seleccione dos equipos de cada uno de los 8 grupos.
									</span>
								</td>
								
								<td class ="centro" width="10%" rowspan="2" style="text-align:center; vertical-align:middle; background-color:#1A497F;">
									<table width="100%">
										<tr>
											<td>
												<a href="#" onclick="Rotar(1)">
													<img src="imagenes/grupos/grp_anterior.png" />
												</a>
											</td>
											
											<td>
												<img id="grpImagen" src="imagenes/banderas/bandera_def.png" />
												<input type="hidden" id="GrupoEdicion" value="" />
											</td>
										</tr>
									</table>
								</td>
							
								<td class ="centro" width="10%" style="text-align:center;">
									<img id="BanderaPos1" src="imagenes/banderas/bandera_def.png" />
									<br />
									<label id="NombreEquipo1"></label>
								</td>
								
								<td class ="centro" width="15%; vertical-align: top;">
									<div id="OpcionBandera01" style="display:none; text-align:left; font-size:14px;">
										<input style="width:20px; height:20px;" id="chkPos01a" type="checkbox" value="1" onclick="ResetPrimeros(this.value)" />1ero
										<br /><br />
										<input style="width:20px; height:20px;" id="chkPos01b" type="checkbox" value="1" onclick="ResetSegundos(this.value)" />2do
									</div>
								</td>
								
								<td class ="centro" style="width:5%;">&nbsp;</td>
								
								<td class ="centro" style="text-align:center; width:10%;">
									<img id="BanderaPos3" src="imagenes/banderas/bandera_def.png" />
									<br />
									<label id="NombreEquipo3"></label>
								</td>
								
								<td class ="centro" style="width=15%; vertical-align: top;">
									<div id="OpcionBandera03" style="display:none; text-align:left; font-size:14px;">				
										<input style="width:20px; height:20px;" id="chkPos03a" type="checkbox" value="3" onclick="ResetPrimeros(this.value)" />1ero
										<br /><br />
										<input style="width:20px; height:20px;" id="chkPos03b" type="checkbox" value="3" onclick="ResetSegundos(this.value)" />2do
									</div>
								</td>
								
								<td class ="centro" width="10%" rowspan="2" style="text-align:center; vertical-align:middle; background-color:#1A497F;">
									<a href="#" onclick="Rotar(2)">
										<img src="imagenes/grupos/grp_siguiente.png" />
									</a>
								</td>					
								
								<td class ="centro" style="width: 15%;">
									<span style="font-size:10px">
										Al finalizar su seleccion, utilize el cuadro en la 
										parte inferior para seleccionar los resultados de cada
									</span>
								</td>					
							</tr>

							<tr style="background-color:#F8F8F8;">
							<td class ="centro" style="width:15%;">
								<span style="font-size:10px">
									Especifique cual de ellos ocupara el primer y segundo lugar de su grupo
								</span>
							</td>

								<td class ="centro" style="text-align:center; width: 10%;">
									<img id="BanderaPos2" src="imagenes/banderas/bandera_def.png" />
									<br />
									<label id="NombreEquipo2"></label>
								</td>
								
								<td class ="centro" width="15%; vertical-align: top;">
									<div id="OpcionBandera02" style="display:none; text-align:left; font-size:14px;">
										<input style="width:20px; height:20px;" id="chkPos02a" type="checkbox" value="2" onclick="ResetPrimeros(this.value)"/>1ero
										<br /><br />
										<input style="width:20px; height:20px;" id="chkPos02b" type="checkbox" value="2" onclick="ResetSegundos(this.value)" />2do
									</div>
								</td>
								
								<td class ="centro" style="width:5%;">&nbsp;</td>
								
								<td class ="centro" style="text-align:center; width:10%;">
									<img id="BanderaPos4" src="imagenes/banderas/bandera_def.png" />
									<br />
									<label id="NombreEquipo4"></label>
								</td>
								
								<td class ="centro" style="width=15%; vertical-align: top;">
									<div id="OpcionBandera04" style="display:none; text-align:left; font-size:14px;">
										<input style="width:20px; height:20px;" id="chkPos04a" type="checkbox" value="4" onclick="ResetPrimeros(this.value)" />1ero
										<br /><br />
										<input style="width:20px; height:20px;" id="chkPos04b" type="checkbox" value="4" onclick="ResetSegundos(this.value)" />2do
									</div>
								</td>
								<td class ="centro" style="width: 15%;">
									<span style="font-size:10px">
										partido, haciendo "click" sobre la bandera del equipo 
										que considera ganador.
									</span>
								</td>
							</tr>				
						</table>

					</div>
					
					<div id="Seleccionados" align="center">
						<table width="100%" cellpadding="8" cellspacing="0" style="background-color:#E0E0E0;">
							<tr>
								<td class ="centro" width="250px" onclick="toggleGrupos(0)" style="background-color:#000000; color:#FFFFFF; font-size:16px;">&nbsp;
									
								</td>
							
								<td class ="centro" width="500px" style="background-color:#000000; color:#FFFFFF; font-size:16px;">
									EQUIPOS SELECCIONADOS
								</td>
								
								<td class ="centro" width="250px" style="background-color:#000000; color:#FFFFFF; font-size:16px;">&nbsp;
									
								</td>					
							</tr>			
						</table>		
					
						<table width="100%" cellspacing="0"style="background-color:#E2E2E2;">
							<tr>
								<td class ="centro" width="48px">&nbsp;</td>
								<td class ="centro" width="25px">&nbsp;</td>
								<td class ="centro" width="50px">&nbsp;</td>
								
								<td class ="centro" width="47px">&nbsp;</td>
								<td class ="centro" width="25px">&nbsp;</td>
								<td class ="centro" width="50px">&nbsp;</td>
								
								<td class ="centro" width="47px">&nbsp;</td>
								<td class ="centro" width="25px">&nbsp;</td>
								<td class ="centro" width="50px">&nbsp;</td>

								<td class ="centro" width="47px">&nbsp;</td>
								<td class ="centro" width="25px">&nbsp;</td>
								<td class ="centro" width="50px">&nbsp;</td>
								<td class ="centro" width="47px">&nbsp;</td>
								<td class ="centro" width="25px">&nbsp;</td>
								<td class ="centro" width="50px">&nbsp;</td>

								<td class ="centro" width="47px">&nbsp;</td>
								<td class ="centro" width="25px">&nbsp;</td>
								<td class ="centro" width="50px">&nbsp;</td>

								<td class ="centro" width="47px">&nbsp;</td>
								<td class ="centro" width="25px">&nbsp;</td>
								<td class ="centro" width="50px">&nbsp;</td>

								<td class ="centro" width="47px">&nbsp;</td>
								<td class ="centro" width="25px">&nbsp;</td>
								<td class ="centro" width="50px">&nbsp;</td>
								
								<td class ="centro" width="48px">&nbsp;</td>
							</tr>
						
							<tr>
								<td>&nbsp;</td>
								<td class ="centro" width="21" rowspan="2">
									<a href="#" onclick="CambiarGrupo('A')" id="grupoA_tab" class="menu">
										<img src="imagenes/grupos/grpA.jpg" width="21" height="85" />
									</a>
								</td>
								<td class ="centro" width="118">
									<img id="BanderaGrpA1" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" />
									<br /><label id="NombreGrpA1"></label>
								</td>
								
								<td>&nbsp;</td>
								<td class ="centro" width="21" rowspan="2">
									<a href="#" onclick="CambiarGrupo('B')" id="grupoB_tab" class="menu">
										<img src="imagenes/grupos/grpB.jpg" width="21" height="85" />
									</a>
								</td>
								<td class ="centro" width="133">
									<img id="BanderaGrpB1" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" />
									<br /><label id="NombreGrpB1"></label>						
								</td>
								
								<td>&nbsp;</td>
								<td class ="centro" width="21" rowspan="2">
									<a href="#" onclick="CambiarGrupo('C')" id="grupoC_tab" class="menu">						
										<img src="imagenes/grupos/grpC.jpg" width="21" height="85" />
									</a>
								</td>
								<td class ="centro" width="115">
									<img id="BanderaGrpC1" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" />
									<br /><label id="NombreGrpC1"></label>						
								</td>
								
								<td>&nbsp;</td>
								<td class ="centro" width="26" rowspan="2">
									<a href="#" onclick="CambiarGrupo('D')" id="grupoD_tab" class="menu">
										<img src="imagenes/grupos/grpD.jpg" width="21" height="85" />
									</a>
								</td>
								<td class ="centro" width="110">
									<img id="BanderaGrpD1" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" />
									<br /><label id="NombreGrpD1"></label>						
								</td>
								
								<td>&nbsp;</td>
								<td class ="centro" width="21" rowspan="2">
									<a href="#" onclick="CambiarGrupo('E')" id="grupoE_tab" class="menu">
										<img src="imagenes/grupos/grpE.jpg" width="21" height="85" />
									</a>
								</td>
								<td class ="centro" width="113">
									<img id="BanderaGrpE1" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" />
									<br /><label id="NombreGrpE1"></label>						
								</td>
								
								<td>&nbsp;</td>
								<td class ="centro" width="21" rowspan="2">
									<a href="#" onclick="CambiarGrupo('F')" id="grupoF_tab" class="menu">					
										<img src="imagenes/grupos/grpF.jpg" width="21" height="85" />
									</a>
								</td>
								<td class ="centro" width="112">
									<img id="BanderaGrpF1" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" />
									<br /><label id="NombreGrpF1"></label>						
								</td>
								
								<td>&nbsp;</td>
								<td class ="centro" width="21" rowspan="2">
									<a href="#" onclick="CambiarGrupo('G')" id="grupoG_tab" class="menu">					
										<img src="imagenes/grupos/grpG.jpg" width="21" height="85" />
									</a>
								</td>
								<td class ="centro" width="105">
									<img id="BanderaGrpG1" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" />
									<br /><label id="NombreGrpG1"></label>
								</td>
								
								<td>&nbsp;</td>
								<td class ="centro" width="21" rowspan="2">
									<a href="#" onclick="CambiarGrupo('H')" id="grupoH_tab" class="menu">					
										<img src="imagenes/grupos/grpH.jpg" width="21" height="85" />
									</a>
								</td>
								<td class ="centro" width="91">
									<img id="BanderaGrpH1" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" />
									<br /><label id="NombreGrpH1"></label>						
								</td>
								
								<td>&nbsp;</td>
							</tr>
							
							<tr>
								<td>&nbsp;</td>
								<td>
									<img id="BanderaGrpA2" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" />
									<br /><label id="NombreGrpA2"></label>
								</td>

								<td>&nbsp;</td>
								<td>
									<img id="BanderaGrpB2" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" />
									<br /><label id="NombreGrpB2"></label>
								</td>
								
								<td>&nbsp;</td>					
								<td>
									<img id="BanderaGrpC2" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" />
									<br /><label id="NombreGrpC2"></label>
								</td>

								<td>&nbsp;</td>
								<td>
									<img id="BanderaGrpD2" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" />
									<br /><label id="NombreGrpD2"></label>
								</td>

								<td>&nbsp;</td>					
								<td>
									<img id="BanderaGrpE2" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" />
									<br /><label id="NombreGrpE2"></label>
								</td>

								<td>&nbsp;</td>					
								<td>
									<img id="BanderaGrpF2" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" />
									<br /><label id="NombreGrpF2"></label>
								</td>
								
								<td>&nbsp;</td>					
								<td>
									<img id="BanderaGrpG2" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" />
									<br /><label id="NombreGrpG2"></label>
								</td>
								
								<td>&nbsp;</td>					
								<td>
									<img id="BanderaGrpH2" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" />
									<br /><label id="NombreGrpH2"></label>
								</td>
								
								<td>&nbsp;</td>					
							</tr>
							
							<tr>
								<td class ="centro" colspan="25">&nbsp;</td>
							</tr>
					</table>
					</div>		
					
					<div id="Mapa" align="center">
						<table style="margin-left: auto; margin-right: auto; width: 100%; background-image:url(Imagenes/estadio3_res.jpg); background-size: 100% 750px; border-style: none;">
							<tr>
								<td class ="centro" width="111">
								<td class ="centro" width="111">
								<td class ="centro" width="111">
								<td class ="centro" width="111">
								<td class ="centro" width="112">					
								<td class ="centro" width="111">
								<td class ="centro" width="111">
								<td class ="centro" width="111">
							<td class ="centro" width="111">
						</tr>
						
							<tr>
								<td class ="centro"  height="80">
									<img id="Fase1Bandera1" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px" onclick="Promover(1,1,2,1)" />


									<br /><label id="Fase1Nombre1" style="color:#FFFFFF; background-color:#000;"></label>
								</td>
								
								<td class ="centro" rowspan="2">
									<img id="Fase2Bandera1" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px" onclick="Promover(2,1,3,1)" />
									<br /><label id="Fase2Nombre1" style="color:#FFFFFF; background-color:#000;"></label>
								</td>
								
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class ="centro" rowspan="2">
									<div id="CopaCampeon" style="text-align:center; display:none; font-family:Arial, Helvetica, sans-serif; font-size:18px;font-weight:bold;color:#000;">
										&nbsp;
									</div>
								</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								
								<td class ="centro" rowspan="2">
									<img id="Fase2Bandera5" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px"  onclick="Promover(2,5,3,3)"  />
									<br /><label id="Fase2Nombre5" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td><img id="Fase1Bandera9" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" onclick="Promover(1,9,2,5)"  /><br />
								<label id="Fase1Nombre9" style="color:#FFFFFF; background-color:#000;"></label>					 </td>
							</tr>

							<tr>
								<td class ="centro" height="80">
									<img id="Fase1Bandera2" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" onclick="Promover(1,2,2,1)" />
									<br /><label id="Fase1Nombre2" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td class ="centro" rowspan="2">
									<img id="Fase3Bandera1" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px"  onclick="Promover(3,1,4,1)" />
									<br /><label id="Fase3Nombre1" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td>&nbsp;</td>
								
								<td>&nbsp;</td>
								
								<td class ="centro" rowspan="2">
									<img id="Fase3Bandera3" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px"  onclick="Promover(3,3,4,2)" />
									<br /><label id="Fase3Nombre3" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td>
									<img id="Fase1Bandera10" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" onclick="Promover(1,10,2,5)" />
									<br /><label id="Fase1Nombre10" style="color:#FFFFFF; background-color:#000;"></label>					</td>
							</tr>

							<tr>
								<td class ="centro" height="80">
									<img id="Fase1Bandera3" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px" onclick="Promover(1,3,2,2)" />
									<br /><label id="Fase1Nombre3" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td class ="centro" rowspan="2">
									<img id="Fase2Bandera2" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px"  onclick="Promover(2,2,3,1)" />
									<br /><label id="Fase2Nombre2" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td>&nbsp;</td>
								
								<td class ="centro" align="center">
									<img id="Fase5Bandera1" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px" /><br />
									<label id="Fase5Nombre1" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td>&nbsp;</td>
								
								<td class ="centro" rowspan="2">
									<img id="Fase2Bandera6" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px"  onclick="Promover(2,6,3,3)"  />
									<br /><label id="Fase2Nombre6" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td>
									<img id="Fase1Bandera11" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px" onclick="Promover(1,11,2,6)" />
									<br /><label id="Fase1Nombre11" style="color:#FFFFFF; background-color:#000;"></label>					</td>
							</tr>

							<tr>
								<td class ="centro" height="80">
									<img id="Fase1Bandera4" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" onclick="Promover(1,4,2,2)" />
									<br /><label id="Fase1Nombre4" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								<td>&nbsp;</td>
								
								<td class ="centro" rowspan="2">
									<img id="Fase4Bandera1" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px"  onclick="Promover(4,1,5,1)" />
									<br /><label id="Fase4Nombre1" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td>&nbsp;</td>
								
								<td class ="centro" rowspan="2">
									<img id="Fase4Bandera2" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px"  onclick="Promover(4,2,5,1)" />
									<br /><label id="Fase4Nombre2" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td>&nbsp;</td>
								<td>
									<img id="Fase1Bandera12" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" onclick="Promover(1,12,2,6)" />
									<br /><label id="Fase1Nombre12" style="color:#FFFFFF; background-color:#000;"></label>					</td>
							</tr>

							<tr>
								<td class ="centro" height="80">
									<img id="Fase1Bandera5" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" onclick="Promover(1,5,2,3)" />
									<br /><label id="Fase1Nombre5" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td class ="centro" rowspan="2">
									<img id="Fase2Bandera3" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px"  onclick="Promover(2,3,3,2)"  />
									<br /><label id="Fase2Nombre3" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								
								<td class ="centro" rowspan="2">
									<img id="Fase2Bandera7" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px"  onclick="Promover(2,7,3,4)" />
									<br /><label id="Fase2Nombre7" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td>
									<img id="Fase1Bandera13" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" onclick="Promover(1,13,2,7)" />
									<br /><label id="Fase1Nombre13" style="color:#FFFFFF; background-color:#000;"></label>					</td>
							</tr>

							<tr>
								<td class ="centro" height="80">
									<img id="Fase1Bandera6" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" onclick="Promover(1,6,2,3)"  />
									<br /><label id="Fase1Nombre6" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td class ="centro" rowspan="2">
									<img id="Fase3Bandera2" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px"  onclick="Promover(3,2,4,1)" />
									<br /><label id="Fase3Nombre2" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td class ="centro" colspan="3" rowspan="3">
									<div id="ExplicacionFinal" style="text-align:center; display:none; font-weight:bold;  font-family:Arial, Helvetica, sans-serif; font-size:16px;color:#000;">					
										&nbsp;
									</div>
								</td>
								
								<td class ="centro" rowspan="2">
									<img id="Fase3Bandera4" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px"  onclick="Promover(3,4,4,2)" />
									<br /><label id="Fase3Nombre4" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td>
									<img id="Fase1Bandera14" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" onclick="Promover(1,14,2,7)" />
									<br /><label id="Fase1Nombre14" style="color:#FFFFFF; background-color:#000;"></label>					</td>
							</tr>

							<tr>
								<td class ="centro" height="80">
									<img id="Fase1Bandera7" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" onclick="Promover(1,7,2,4)"  />
									<br /><label id="Fase1Nombre7" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td class ="centro" rowspan="2">
									<img id="Fase2Bandera4" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px"  onclick="Promover(2,4,3,2)" />
									<br /><label id="Fase2Nombre4" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td class ="centro" rowspan="2">
									<img id="Fase2Bandera8" src="imagenes/banderas/bandera_def.png"  width="50px" height="34px" onclick="Promover(2,8,3,4)" />
									<br /><label id="Fase2Nombre8" style="color:#FFFFFF; background-color:#000;"></label>					</td>
								
								<td>
									<img id="Fase1Bandera15" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" onclick="Promover(1,15,2,8)" />
									<br /><label id="Fase1Nombre15" style="color:#FFFFFF; background-color:#000;"></label>					</td>
							</tr>

							<tr>
								<td class ="centro" height="80">
									<img id="Fase1Bandera8" src="imagenes/banderas/bandera_def.png" width="50px" height="34px" onclick="Promover(1,8,2,4)" />
									<br /><label id="Fase1Nombre8" style="color:#FFFFFF; background-color:#000;"></label>				  </td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>
									<img id="Fase1Bandera16" src="imagenes/banderas/bandera_def.png" width="50px" height="34px"  onclick="Promover(1,16,2,8)" />
									<br /><label id="Fase1Nombre16" style="color:#FFFFFF; background-color:#000;"></label>					</td>
							</tr>
							
							<tr><td class ="centro" colspan="9">&nbsp;</td></tr>
						</table>
					</div>  		
					
					<div id="Apuesta" style="display:none;">
						<form name="FormularioDatos" id="FormularioDatos" action="grabar_master.asp" method="post">
							<%
								for k = 1 to 31
									Response.write "<input name='Polla" & k & "' id='Polla" & k & "' value='' />"
								next
							%>
						</form>
					</div>
				</tr>
			</td>
		</table>
		
		<%
			sqlString = "SELECT * FROM mundial_ListaEquipos ORDER BY Grupo;"
						
			set con = server.CreateObject("ADODB.Connection")
			con.open Application("Conn")
			set t = con.execute(sqlString)

			indice = 1
		%>
		
		<div style="display:none;">
			<!--
				Esta seccion crea una tabla mediante ASP.
				
				Las rutinas de JavaScript utilizaran esta tabla
				como una tabla de base de datos, actualizandola
				a medida que el usuario hace selecciones
			-->
		
			<%
				do
			%>
				<input id='Grupo<%= indice %>'   value='<%= t("Grupo") %>' />
				<input id='Codigo<%= indice %>'  value='<%= t("Equipo") %>' />
				<input id='Nombre<%= indice %>'  value='<%= t("Nombre") %>' />
				<input id='Bandera<%= indice %>' value='<%= t("Imagen") %>' />
				<input id='Posicion<%= indice %>' value='' />
			<%
					indice = indice + 1
					t.movenext
				loop until t.eof
				
				t.close: set t =nothing				
				
				sqlString = "SELECT * FROM mundial_Master ORDER BY IndiceCuadro;"
				set t = con.execute(sqlString)
	
				indice = 1
				
				do
			%>
				<input id='MasterIndiceCuadro<%= indice %>'   value='<%= t("IndiceCuadro") %>' />
				<input id='MasterFase<%= indice %>'  value='<%= t("Fase") %>' />
				<input id='MasterPosicionGrupos<%= indice %>'  value='<%= t("PosicionGrupos") %>' />
				<input id='MasterCodigoEquipo<%= indice %>' value='<%= t("CodigoEquipo") %>' />
			<%
					indice = indice + 1
					t.movenext
				Loop Until t.eof				
			%>
		</div>
		
		<%
			t.close: set t =nothing
			con.close: set con = nothing
		%>	

		<br /><br /><br />	

		<script type="text/javascript">
			ToggleSW = 1

            function Requery() {
                document.getElementById("formulario").submit();
            } 
		
			function CambiarGrupo(grupo) {
				var indice, campo, valorCampo, bandera, currentGrupo;

				indice = 1;
				document.getElementById("grpImagen").src = 'imagenes/grupos/grp' + grupo + ".jpg";
				document.getElementById("GrupoEdicion").value = grupo;

                removeClass('.menu');
                addClass('#grupo' + grupo);
				
				for (var i = 1; i <= 32; i++) { 
					campo = "Grupo" + i;
					valorCampo = document.getElementById(campo).value;
					
					if(valorCampo == grupo) {
						document.getElementById("BanderaPos" + indice).src = 'imagenes/banderas/' + document.getElementById("Bandera" + i).value;
						document.getElementById('NombreEquipo' + indice).innerHTML = document.getElementById("Nombre" + i).value;
						
						//
						// Verificamos su valor en la tabla actual...
						//
						document.getElementById("OpcionBandera0" + indice).style.display = "block";

						document.getElementById("chkPos0" + indice + "a").checked = false;
						document.getElementById("chkPos0" + indice + "b").checked = false;
						
						if(document.getElementById("Posicion" + i).value == '1')  {
							document.getElementById("chkPos0" + indice + "a").checked = true;
						}

						if(document.getElementById("Posicion" + i).value == '2')  {
							document.getElementById("chkPos0" + indice + "b").checked = true;
						}
								
						indice+=1 ;
					}
				}
			}
			
			function limpiar(cuales, indice) {
				for (var i = 1; i <= 4; i++) { 
					document.getElementById("chkPos0" + i + cuales).checked = false;
				}
			}
			
			function limpiarEquipos() {
				for (var i = 1; i <= 4; i++) { 
					ActualizarPosicionEquipo(document.getElementById("NombreEquipo" + i).innerHTML, "")
				}
			}
			
			function RevisarEqipos() {
				for (var i = 1; i <= 4; i++) { 
					if(document.getElementById("chkPos0" + i + "a").checked == true) {
						ActualizarPosicionEquipo(document.getElementById("NombreEquipo" + i).innerHTML, "1")
					}

					if(document.getElementById("chkPos0" + i + "b").checked == true) {
						ActualizarPosicionEquipo(document.getElementById("NombreEquipo" + i).innerHTML, "2")
					}
				}
			}
			
			function ResetPrimeros(indice) {
				var NombreObsoleto, BanderaObsoleta;
				var NombreNuevo, BanderaNueva;
			
				if(document.getElementById("chkPos0" + indice + "b").checked == true) {
					alert("El mismo equipo no puede ocupar ambas posiciones!");
					document.getElementById("chkPos0" + indice + "a").checked = false;
				}
				else {
					limpiar("a", indice);
					limpiarEquipos();
				
					document.getElementById("chkPos0" + indice + "a").checked = true;
					ActualizarPosicionEquipo(document.getElementById("NombreEquipo" + indice).innerHTML, "1")
					
					RevisarEqipos();
					
					//
					// Antes de Actualizar la barra de seleccion, necesitamos saber
					// cual equipo reemplaza al ya existia antes, de otra forma
					// no podemos actualizar el mapa
					//
					NombreObsoleto = document.getElementById("NombreGrp" + document.getElementById("GrupoEdicion").value + "1").innerHTML;
					BanderaObsoleta = document.getElementById("BanderaGrp" + document.getElementById("GrupoEdicion").value + "1").src;
					
					NombreNuevo = document.getElementById("NombreEquipo" + indice).innerHTML;
					BanderaNueva = document.getElementById("BanderaPos" + indice).src;
					
					ActualizarSeleccion();
					PresentarBanderas();	
					
					if(NombreObsoleto!="") {
						RevisarPromociones(NombreObsoleto, BanderaObsoleta, NombreNuevo, BanderaNueva);
					}
	
					Completo();
				}
			}
			
			function ResetSegundos(indice) {
				var NombreObsoleto, BanderaObsoleta;
							
				if (document.getElementById("chkPos0" + indice + "a").checked == true) {
					alert("El mismo equipo no puede ocupar ambas posiciones!");		
					document.getElementById("chkPos0" + indice + "b").checked = false;
				}
				else {
					limpiar("b", indice);
					limpiarEquipos();
				
					document.getElementById("chkPos0" + indice + "b").checked = true;	
					ActualizarPosicionEquipo(document.getElementById("NombreEquipo" + indice).innerHTML, "2")
					
					RevisarEqipos();
					
					NombreObsoleto = document.getElementById("NombreGrp" + document.getElementById("GrupoEdicion").value + "2").innerHTML;
					BanderaObsoleta = document.getElementById("BanderaGrp" + document.getElementById("GrupoEdicion").value + "2").src;
					
					NombreNuevo = document.getElementById("NombreEquipo" + indice).innerHTML;
					BanderaNueva = document.getElementById("BanderaPos" + indice).src;
					
					ActualizarSeleccion();
					PresentarBanderas();
					
					if(NombreObsoleto!="") {
						RevisarPromociones(NombreObsoleto, BanderaObsoleta, NombreNuevo, BanderaNueva);
					}
	
					Completo();				
				}
			}
			
			function ActualizarPosicionEquipo(NombreEquipo, Posicion) {
				for (var i = 1; i <= 32; i++) { 
					if(document.getElementById("Nombre" + i).value == NombreEquipo) {
						document.getElementById("Posicion" + i).value = Posicion;
					}
				}				
			}

			function ActualizarPosicionCodigoEquipo(CodigoEquipo, Posicion) {
				for (var i = 1; i <= 32; i++) { 
					if(document.getElementById("Codigo" + i).value == CodigoEquipo) {
						document.getElementById("Posicion" + i).value = Posicion;
					}
				}				
			}
			
			function ActualizarSeleccion() {
				for (var i = 1; i <= 32; i++) { 
					if(document.getElementById("Posicion" + i).value != "") {
						campo = "BanderaGrp" + document.getElementById("Grupo" + i).value + document.getElementById("Posicion" + i).value;
						document.getElementById(campo).src = 'imagenes/banderas/' + document.getElementById("Bandera" + i).value;
						
						campo = "NombreGrp" + document.getElementById("Grupo" + i).value + document.getElementById("Posicion" + i).value;
						document.getElementById(campo).innerHTML = document.getElementById("Nombre" + i).value;							
					}
				}
			}
			
			function Presentar() {
				//
				// Verificamos que se hayan seleccionado 16 Equipos...
				//
				var contador = 0, campo;
				
				for (var i = 1; i <= 32; i++) { 
					if(document.getElementById("Posicion" + i).value != "") contador+=1;
				}
				
				if (contador == 16) {
					document.getElementById("Preliminares").style.display = "none";				
					document.getElementById("Mapa").style.display = "block";
				
					for (var i = 1; i <= 32; i++) { 
						if(document.getElementById("Posicion" + i).value != "") {
							PresentarBanderas();
						}
					}					
				}
				else {
					alert("No se han seleccionado todos los equipos necesarios! Por favor, verifique y vuelva a intentarlo.");
				}
			}
			
			function PresentarBanderas() {
				document.getElementById("Fase1Bandera1").src = document.getElementById("BanderaGrpA1").src;
				document.getElementById("Fase1Nombre1").innerHTML = document.getElementById("NombreGrpA1").innerHTML;

				document.getElementById("Fase1Bandera2").src = document.getElementById("BanderaGrpB2").src;
				document.getElementById("Fase1Nombre2").innerHTML = document.getElementById("NombreGrpB2").innerHTML;

				document.getElementById("Fase1Bandera3").src = document.getElementById("BanderaGrpC1").src;
				document.getElementById("Fase1Nombre3").innerHTML = document.getElementById("NombreGrpC1").innerHTML;

				document.getElementById("Fase1Bandera4").src = document.getElementById("BanderaGrpD2").src;
				document.getElementById("Fase1Nombre4").innerHTML = document.getElementById("NombreGrpD2").innerHTML;

				document.getElementById("Fase1Bandera5").src = document.getElementById("BanderaGrpE1").src;
				document.getElementById("Fase1Nombre5").innerHTML = document.getElementById("NombreGrpE1").innerHTML;

				document.getElementById("Fase1Bandera6").src = document.getElementById("BanderaGrpF2").src;
				document.getElementById("Fase1Nombre6").innerHTML = document.getElementById("NombreGrpF2").innerHTML;

				document.getElementById("Fase1Bandera7").src = document.getElementById("BanderaGrpG1").src;
				document.getElementById("Fase1Nombre7").innerHTML = document.getElementById("NombreGrpG1").innerHTML;

				document.getElementById("Fase1Bandera8").src = document.getElementById("BanderaGrpH2").src;
				document.getElementById("Fase1Nombre8").innerHTML = document.getElementById("NombreGrpH2").innerHTML;
				
				document.getElementById("Fase1Bandera9").src = document.getElementById("BanderaGrpB1").src;
				document.getElementById("Fase1Nombre9").innerHTML = document.getElementById("NombreGrpB1").innerHTML;
				
				document.getElementById("Fase1Bandera10").src = document.getElementById("BanderaGrpA2").src;
				document.getElementById("Fase1Nombre10").innerHTML = document.getElementById("NombreGrpA2").innerHTML;
				
				document.getElementById("Fase1Bandera11").src = document.getElementById("BanderaGrpD1").src;
				document.getElementById("Fase1Nombre11").innerHTML = document.getElementById("NombreGrpD1").innerHTML;
				
				document.getElementById("Fase1Bandera12").src = document.getElementById("BanderaGrpC2").src;
				document.getElementById("Fase1Nombre12").innerHTML = document.getElementById("NombreGrpC2").innerHTML;
				
				document.getElementById("Fase1Bandera13").src = document.getElementById("BanderaGrpF1").src;
				document.getElementById("Fase1Nombre13").innerHTML = document.getElementById("NombreGrpF1").innerHTML;
				
				document.getElementById("Fase1Bandera14").src = document.getElementById("BanderaGrpE2").src;
				document.getElementById("Fase1Nombre14").innerHTML = document.getElementById("NombreGrpE2").innerHTML;
				
				document.getElementById("Fase1Bandera15").src = document.getElementById("BanderaGrpH1").src;
				document.getElementById("Fase1Nombre15").innerHTML = document.getElementById("NombreGrpH1").innerHTML;
				
				document.getElementById("Fase1Bandera16").src = document.getElementById("BanderaGrpG2").src;
				document.getElementById("Fase1Nombre16").innerHTML = document.getElementById("NombreGrpG2").innerHTML;			
			}

			function RevisarPromociones(NombreObsoleto, BanderaObsoleta, NombreNuevo, BanderaNueva) {
				var limite=32;
		
				for(fase=1; fase<=5; fase++) {
					limite/=2;
				
					for(var equipo=1; equipo<=limite; equipo++) {
						if(document.getElementById("Fase" + fase + "Nombre" + equipo).innerHTML == NombreObsoleto) {
							document.getElementById("Fase" + fase + "Nombre" + equipo).innerHTML = NombreNuevo;
							document.getElementById("Fase" + fase + "Bandera" + equipo).src = BanderaNueva;
						}
						
					}
				}
			}

			function Promover(a,b,c,d) {
				var fuente1 = document.getElementById("Fase" + a + "Nombre" + b).innerHTML;
				var fuente2 = document.getElementById("Fase" + a + "Bandera" + b).src;
				var destino1 = document.getElementById("Fase" + c + "Nombre" + d).innerHTML;
				
				document.getElementById("Fase" + c + "Nombre" + d).innerHTML = fuente1;
				document.getElementById("Fase" + c + "Bandera" + d).src = fuente2;
				
				if (destino1 != "") {		
					ReemplazoGlobal(fuente1, fuente2, destino1, a);
				}
				
				if(c==5 || d==1) {
					if(document.getElementById("Fase5Nombre1").innerHTML != "") {
						document.getElementById("CopaCampeon").style.display="block";
						document.getElementById("ExplicacionFinal").style.display="block";						
					}
				}
				
				Completo();
			}
			
			function ReemplazoGlobal(NuevoTexto, NuevaBandera, ViejoTexto, FaseInicial) {
				var limite;
				var objeto;
				
				switch (FaseInicial) {
					case 1:
						limite = 16;
						break;
					case 2:
						limite = 8;					
						break;
					case 3:
						limite = 4;
						break;
					case 4:
						limite = 2;					
						break;
				}				
				
				if(FaseInicial<5) {
					for(var fase=(FaseInicial + 1); fase<=5; fase++) {

						limite/=2;
						
						for(var equipo=1; equipo<=limite; equipo++) {
							objeto = document.getElementById("Fase" + fase + "Nombre" + equipo).innerHTML;
							
							if(objeto==ViejoTexto) {
								//
								// Este objeto debe ser reemplazado
								//
								document.getElementById("Fase" + fase + "Nombre" + equipo).innerHTML = NuevoTexto;
								document.getElementById("Fase" + fase + "Bandera" + equipo).src = NuevaBandera;
							}
						}	
					}
				}
			}
			
			function toggleGrupos(opcion) {
				if (ToggleSW == 1) {
					document.getElementById("Preliminares").style.display="none";
					document.getElementById("Seleccionados").style.display="none";
					document.getElementById("toggleOcultarGrupos").innerHTML="VER GRUPOS";
					document.getElementById("SeleccionGrupo").style.display="none";
					document.getElementById("ActualizarEquipos").style.display="block";					
				}
				else {
					document.getElementById("Preliminares").style.display="block";
					document.getElementById("Seleccionados").style.display="block";									
					document.getElementById("toggleOcultarGrupos").innerHTML="OCULTAR GRUPOS";
					document.getElementById("SeleccionGrupo").style.display="block";
					document.getElementById("ActualizarEquipos").style.display="none";
				}
				
				if(ToggleSW == 1) ToggleSW = 0
				else ToggleSW = 1
			}

			function Completo() {
				var c = 0;
				var e = 32;
				
				for(j=1;j<=5;j++) {
					e/=2;
					
					for(k=1;k<=e;k++) {
						if(document.getElementById("Fase" + j + "Nombre" + k).innerHTML != "") {
							c+=1;
						}
					}
				}
			}

			function Rotar(direccion) {
				var Grupos = new Array("", "A", "B", "C", "D", "E", "F", "G", "H"); 
				var k, donde;
				
				donde = 0;
				
				for(k=1;k<=8;k++) {
					if(Grupos[k] == document.getElementById("GrupoEdicion").value) {
						donde = k;
					}
				}
				
				if (direccion == 1) {
					//izquierda
					if(donde == 1) donde = 8
					else donde -=1
				}
				else {
					// derecha
					if(donde == 8) donde = 1
					else donde += 1
				}
				
				CambiarGrupo(Grupos[donde]);			
			}

			function PasarSubmit() {
				var limite = 32;
				var indice = 1;
				var NombreEquipo = "";
				var Codigo = "";
				
				for(var fase=1; fase<=5; fase++) {
					limite/=2;
					
					for(var equipos=1;equipos<=limite;equipos++) {
						NombreEquipo = document.getElementById("Fase" + fase + "Nombre" + equipos).innerHTML;
						Codigo = CodigoEquipo(NombreEquipo);
						document.getElementById("Polla" + indice).value = Codigo;
						indice+=1;
					}
				}
				
				document.FormularioDatos.submit();
			}
			
			function CodigoEquipo(NombreEquipo) {
				var res = '';
				
				for (var k=1; k <=32; k++) {
					if (document.getElementById("Nombre" + k).value == NombreEquipo) {
						res = document.getElementById("Codigo" + k).value;
					}
				}
				
				return res;
			}

			function NombreEquipo(CodigoEquipo) {
				var nombre=""; 
				
				for(var k=1; k<=32;k++) {
					if(document.getElementById("Codigo" + k).value == CodigoEquipo) {
						nombre = document.getElementById("Nombre" + k).value;
					}
				}
				return nombre;
			}

			function NombreBandera(CodigoEquipo) {
				var nombre=""; 
				
				for(var k=1; k<=32;k++) {
					if(document.getElementById("Codigo" + k).value == CodigoEquipo) {
						nombre = document.getElementById("Bandera" + k).value;
					}
				}
				return nombre;
			}

			function NombreEquipoPosGrp(PosGrupo) {
				var nombre=""; 
				
				for(var k=1; k<=31;k++) {
					if(document.getElementById("MasterPosicionGrupos" + k).value == PosGrupo) {
						nombre = NombreEquipo(document.getElementById("MasterCodigoEquipo" + k).value);
					}
				}
				return nombre;
			}

			function NombreBanderaPosGrp(PosGrupo) {
				var nombre=""; 
				
				for(var k=1; k<=31;k++) {
					if(document.getElementById("MasterPosicionGrupos" + k).value == PosGrupo) {
						nombre = NombreBandera(document.getElementById("MasterCodigoEquipo" + k).value);
					}
				}
				return nombre;
			}

			function CodigoPosGrp(PosGrupo) {
				var nombre=""; 
				
				for(var k=1; k<=31;k++) {
					if(document.getElementById("MasterPosicionGrupos" + k).value == PosGrupo) {
						nombre = document.getElementById("MasterCodigoEquipo" + k).value;
					}
				}
				return nombre;
			}

			function PresentarMaster() {
				//
				// Actualizamos la Seleccion de Equipos...
				//		
				
				//ActualizarPosicionCodigoEquipo(CodigoEquipo, Posicion)
				
				var grupos = new Array("", "A", "B", "C", "D", "E", "F", "G", "H");
				var nomGrupo = "";

				//
				// Actualizamos los Selectores de Grupos...
				//					

				for(g=1;g<=8;g++) {
					for(k=1; k<=2; k++) {
						nomGrupo = grupos[g] + k;
						
						if(k==1) { ActualizarPosicionCodigoEquipo(CodigoPosGrp(nomGrupo), 1); }
						else {
							if(k==2) { ActualizarPosicionCodigoEquipo(CodigoPosGrp(nomGrupo), 2); }
						}
					}
				}	


				//
				// Actualizamos los grupos...
				//		
				
			
				for(g=1;g<=8;g++) {
					for(k=1; k<=2; k++) {
						nomGrupo = grupos[g] + k;

						if(NombreBanderaPosGrp(nomGrupo) != '') {
							document.getElementById("BanderaGrp" + nomGrupo).src = 'imagenes/banderas/' + NombreBanderaPosGrp(nomGrupo);
							document.getElementById("NombreGrp" + nomGrupo).innerHTML = NombreEquipoPosGrp(nomGrupo);
						}
					}
				}				
				
				//
				// Actualizamos el Cuadro de Juegos...
				//
				
				var indiceMaster = 1;
				var bandera = "";
				var limite = 32;				
				
				for(var fase=1;fase<=5;fase++) {
					limite/=2;
					
					for(var equipo=1;equipo<=limite;equipo++) {
						bandera = "imagenes/banderas/";
						
						if(document.getElementById("MasterCodigoEquipo" + indiceMaster).value!="-") {
							document.getElementById("Fase" + fase + "Nombre" + equipo).innerHTML = NombreEquipo(document.getElementById("MasterCodigoEquipo" + indiceMaster).value);
							document.getElementById("Fase" + fase + "Bandera" + equipo).src = 'imagenes/banderas/' +  NombreBandera(document.getElementById("MasterCodigoEquipo" + indiceMaster).value);						
						}
						indiceMaster++;
					}
				}
			}

            function removeClass(objeto, clase = 'active') {
                if (typeof objeto === 'string') {
                    document.querySelectorAll(objeto).forEach(el => el.classList.remove(clase));
                } else if (objeto instanceof Element) {
                    objeto.classList.remove(clase);
                }
            }

            function addClass(objeto, clase = 'active') {
                if (typeof objeto === 'string') {
                    document.querySelectorAll(objeto).forEach(el => el.classList.add(clase));
                } else if (objeto instanceof Element) {
                    objeto.classList.add(clase);
                }
            }            
		</script>	
		<!-- #include virtual = "/core/includes/kernel/close.inc" -->	
	</body>
</html>