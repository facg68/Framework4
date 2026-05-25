<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
		<title>Polla Mundial</title>
		
		<style type="text/css">
			td {
				font-family:Verdana, Arial, Helvetica, sans-serif;
				font-size:14px;
				text-align:center;
			}
			
			a.menu:link 	{ font-family:Verdana, Arial, Helvetica, sans-serif; font-size:16px; color:#FFF; text-decoration: none; }
			a.menu:visited 	{ font-family:Verdana, Arial, Helvetica, sans-serif; font-size:16px; color:#FFF; text-decoration: none; }
			a.menu:hover 	{ font-family:Verdana, Arial, Helvetica, sans-serif; font-size:18px; color:#FFFF4F; font-weight:bold;}
		</style>

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
	
	<body onload="Presentar()">
		<div style="display:none;">	
			<%
				Dim Tiquete, k, campo, con, sqlString, t
				Dim Nombre, Estatus, FechaConfeccion
				
				Tiquete = Request.QueryString("t")
				If Tiquete = "" then
					Tiquete = Request.Form("txtCodigoUnico")
				end if
				
				sqlString = "Exec dbo.mundial_PollaPuntos '" & Tiquete & "';"
				
				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)
				
				if not(t.bof or t.eof) then
					Nombre = t("Nombre")
					Estatus = t("Estatus")
					FechaConfeccion = t("FechaConfeccion")
					
					do
						%>
							<input id='NombreEquipo<%= t("Posicion") %>'	value='<%= t("NombreEquipo") %>' />	
							<input id='Bandera<%= t("Posicion") %>'			value='<%= t("Imagen") %>' />
						<%
						t.movenext
					loop until (t.eof)
				end if
				
				t.close: set t =nothing
				con.close: set con = nothing
			%>		
		</div>
		
		<table width="1000px" align="center" cellpadding="0" cellspacing="0">
			<tr>
				<td>
					<img src="Imagenes/header.jpg" style="width:100%; border-style: none;" onclick="Back()">
				</td>
			</tr>
		</table>
		
		<table width="1000px" align="center" border="0" cellpadding="8" cellspacing="0">
			<tr>
				<td colspan="2" style="font-family:Verdana, Arial, Helvetica, sans-serif; font-size:20px; font-weight:bold; text-align:center;">
					<%= "BOLETO  " & Tiquete %>
				</td>
			</tr>
		
			<tr>
				<td width="50%" style="text-align:left;">
					<strong>APUESTA DE </strong><%= UCase(Nombre) %><br />
					<strong>FECHA DE CONFECCION </strong><%= FechaConfeccion %><br />
					<strong>ESTATUS </strong>
					<%
						if Estatus = 1 then
							response.Write "Activado"
						else
							response.write "Sin Activar"
						end if
					%>
				</td>
				
				<td>
					Firma de Informatica: _____________________________
					<br />
					<br />
					Firma del Usuario: ___________________________
				</td>
			</tr>
		</table>		
		
			<table align="center" width="1000px" style="background-image:url(Imagenes/estadio3.jpg);">
				<tr>
					<td width="111">
					<td width="111">
					<td width="111">
					<td width="111">
					<td width="112">					
					<td width="111">
					<td width="111">
					<td width="111">
					<td width="111">
				</tr>
			
				<tr>
					<td  height="80">
						<img id="Fase1Bandera1" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" />
						<br /><label id="Fase1Nombre1" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td rowspan="2">
						<img id="Fase2Bandera1" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px"/>
						<br /><label id="Fase2Nombre1" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					
					<td rowspan="2" style="vertical-align:bottom;">
						<img src="Imagenes/copa.png" border="0" /><br />
						CAMPEON					
					</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					
					<td rowspan="2">
						<img id="Fase2Bandera5" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" />
						<br /><label id="Fase2Nombre5" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td><img id="Fase1Bandera9" src="imagenes/banderas/bandera_def.png" border="0" width="50px" height="34px" /><br />
					  <label id="Fase1Nombre9" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
				</tr>

				<tr>
					<td height="80">
						<img id="Fase1Bandera2" src="imagenes/banderas/bandera_def.png" border="0" width="50px" height="34px" />
						<br /><label id="Fase1Nombre2" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td rowspan="2">
						<img id="Fase3Bandera1" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" />
						<br /><label id="Fase3Nombre1" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td>&nbsp;</td>
					
					<td>&nbsp;</td>
					
					<td rowspan="2">
						<img id="Fase3Bandera3" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" />
						<br /><label id="Fase3Nombre3" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td>
						<img id="Fase1Bandera10" src="imagenes/banderas/bandera_def.png" border="0" width="50px" height="34px" />
						<br /><label id="Fase1Nombre10" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
				</tr>

				<tr>
					<td height="80">
						<img id="Fase1Bandera3" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" />
						<br /><label id="Fase1Nombre3" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td rowspan="2">
						<img id="Fase2Bandera2" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" />
						<br /><label id="Fase2Nombre2" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td>&nbsp;</td>
					
					<td align="center">
						<img id="Fase5Bandera1" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" /><br />
						<label id="Fase5Nombre1" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td>&nbsp;</td>
					
					<td rowspan="2">
						<img id="Fase2Bandera6" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" />
						<br /><label id="Fase2Nombre6" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td>
						<img id="Fase1Bandera11" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" />
						<br /><label id="Fase1Nombre11" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
				</tr>

				<tr>
					<td height="80">
						<img id="Fase1Bandera4" src="imagenes/banderas/bandera_def.png" border="0" width="50px" height="34px" />
						<br /><label id="Fase1Nombre4" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td>&nbsp;</td>
					
					<td rowspan="2">
						<img id="Fase4Bandera1" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" />
						<br /><label id="Fase4Nombre1" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td>&nbsp;</td>
					
					<td rowspan="2">
						<img id="Fase4Bandera2" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" />
						<br /><label id="Fase4Nombre2" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td>&nbsp;</td>
					
					<td>
						<img id="Fase1Bandera12" src="imagenes/banderas/bandera_def.png" border="0" width="50px" height="34px" />
						<br /><label id="Fase1Nombre12" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
				</tr>

				<tr>
					<td height="80">
						<img id="Fase1Bandera5" src="imagenes/banderas/bandera_def.png" border="0" width="50px" height="34px" />
						<br /><label id="Fase1Nombre5" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td rowspan="2">
						<img id="Fase2Bandera3" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" />
						<br /><label id="Fase2Nombre3" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					
					<td rowspan="2">
						<img id="Fase2Bandera7" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" />
						<br /><label id="Fase2Nombre7" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td>
						<img id="Fase1Bandera13" src="imagenes/banderas/bandera_def.png" border="0" width="50px" height="34px" />
						<br /><label id="Fase1Nombre13" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
				</tr>

				<tr>
					<td height="80">
						<img id="Fase1Bandera6" src="imagenes/banderas/bandera_def.png" border="0" width="50px" height="34px" />
						<br /><label id="Fase1Nombre6" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td rowspan="2">
						<img id="Fase3Bandera2" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" />
						<br /><label id="Fase3Nombre2" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td colspan="3" rowspan="3">&nbsp;</td>
					
					<td rowspan="2">
						<img id="Fase3Bandera4" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" />
						<br /><label id="Fase3Nombre4" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td>
						<img id="Fase1Bandera14" src="imagenes/banderas/bandera_def.png" border="0" width="50px" height="34px" />
						<br /><label id="Fase1Nombre14" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
				</tr>

				<tr>
					<td height="80">
						<img id="Fase1Bandera7" src="imagenes/banderas/bandera_def.png" border="0" width="50px" height="34px" />
						<br /><label id="Fase1Nombre7" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td rowspan="2">
						<img id="Fase2Bandera4" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" />
						<br /><label id="Fase2Nombre4" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td rowspan="2">
						<img id="Fase2Bandera8" src="imagenes/banderas/bandera_def.png" border="0"  width="50px" height="34px" />
						<br /><label id="Fase2Nombre8" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td>
						<img id="Fase1Bandera15" src="imagenes/banderas/bandera_def.png" border="0" width="50px" height="34px" />
						<br /><label id="Fase1Nombre15" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
				</tr>

				<tr>
					<td height="80">
						<img id="Fase1Bandera8" src="imagenes/banderas/bandera_def.png" border="0" width="50px" height="34px" />
						<br /><label id="Fase1Nombre8" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
					
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					
					<td>
						<img id="Fase1Bandera16" src="imagenes/banderas/bandera_def.png" border="0" width="50px" height="34px" />
						<br /><label id="Fase1Nombre16" style="color:#FFFFFF; background-color:#000;"></label>
					</td>
				</tr>
				
				<tr><td colspan="9">&nbsp;</td></tr>
			</table>
			
			<table width="1000px" border="0" align="center">
				<tr>
					<td align="center">
						Usted debe Conservar este boleto para reclamar el/los premios al final del juego.<br />
						Si el boleto se encuentra adulterado quedar&aacute; desacalificado autom&aacute;ticamente.
					</td>
				</tr>
			</table>
			
		<form name="Regresar" method="post" action="activar.asp">			
			<input name="txtUsuarioValidado" type="hidden" value="1" />
		</form>

		<script type="text/javascript">
			function Presentar() {
				var campo="";
				var limite = 32;
				var indice = 1;
				
				for(fase=1;fase<=5;fase++) {
					limite/=2;
					
					for(equipo=1;equipo<=limite;equipo++) {
						campo = "NombreEquipo" + indice;
						
						document.getElementById("Fase" + fase + "Nombre" + equipo).innerHTML = document.getElementById("NombreEquipo" + indice).value;
						document.getElementById("Fase" + fase + "Bandera" + equipo).src = "imagenes/banderas/" + document.getElementById("Bandera" + indice).value;
						indice+=1;
					}
				}
			}
			
			function Back() {
				document.Regresar.submit();
			}
		</script>		
	</body>
</html>