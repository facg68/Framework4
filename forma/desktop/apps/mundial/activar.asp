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
            thisProcess = "mundial.040"
            SysLockOut
        %>     
		
        <style>
            body {
                background-color:#9F9F9F;
                background-image:url(Imagenes/fondo.jpg);
            }
			
            td {
                font-family:Verdana, Arial, Helvetica, sans-serif;
                font-size:12px;
            }

			tr:not(:last-child) { border: none !important; }
			
            #content h2 { text-align: left; }

			.fondo0 { background-color:rgb(255, 255, 255); }
			.fondo1 { background-color: rgb(232, 250, 227); }
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
		
	
			Function SePuedeActivar()
				dim con, t, sqlString
				
				sqlString = "SELECT Estatus FROM dbo.mundial_Estatus WHERE Codigo = 'Activar';"
				
				set con = Server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)

				SePuedeActivar = t("Estatus")
				
				t.close: set t=nothing
				con.close: set con=nothing			
			End Function
		%>
    </head>

    <body>
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

		<%
			dim k, campo, con, sqlString, t, indicefondo, multiplicador
			dim ver, validado, contador, ProcedimientoAlmacenado

			ver = request.QueryString("v")
			if ver = "" then ver = 0
			
			set con = Server.CreateObject("ADODB.Connection")
			con.open Application("Conn")

			if (Usuario_Valido() = 0) then
				response.Redirect "mundial.asp"
			else
		%>

			<table style="width: 95%; margin-left: auto; margin-right: auto; padding: 0px; border-spacing: 0px; background-color: rgb(0, 0, 0);">
				<tr style="background-color:rgb(0, 0, 0); color:rgb(255, 255, 255);">
					<td colspan="8">
						<img src="Imagenes/header.jpg" style="width: 100%;">
					</td>
				</tr>


				<tr style="color:rgb(255, 255, 255);">
					<td colspan="8">
						<table style="width: 100%">
							<tr>
								<td style="width: 20%; text-align: left; padding: 10px; font-size:14px;">
									<%
										if ver = 0 then
											response.write "<a style='color: white;' href='activar.asp?v=1&u=1'>Ver Pollas<br />Activadas</a>"
										else
											response.write "<a style='color: white;' href='activar.asp?v=0&u=1'>Ver Pollas<br />Sin Activar</a>"
										end if
									%>
								</td>				
							
								<td style="width: 60%; text-align: center; padding: 10px; font-size:16px; font-weight:bold;">
									<%
										if SePuedeActivar() = 0 then
											response.write "Periodo de Activaciones Finalizado"
										end if
										
										response.write "<br>"
										
										if ver = 0 then
											response.write "Apuestas Sin Activar"
										else
											response.write "Apuestas Activadas"
										end if
									%>
								</td>
							
								<td style="width: 20%; text-align: right; padding: 10px;">
									<a href="default.asp">
										<img src="Imagenes/home.png" style="width: 40px; height: 40px; " />
									</a>
								</td>
							</tr>
						</table>
					</td>
				</tr>

				<!-- 
					Aqui van los campos del formulario... 
				-->

				<%
					if ver = 0 then
						ProcedimientoAlmacenado = "mundial_ApuestasNoActivadas"
					else
						ProcedimientoAlmacenado = "mundial_ApuestasActivadas"
					end if
				
					set t = con.execute(ProcedimientoAlmacenado)
						
					if (t.bof or t.eof) then
						%>
							<tr>
								<td colspan="8" style="text-align: center; padding: 10px;">
									No hay Apuestas Sin Activar en este momento!
								</td>
							</tr>
						<%
					else
						indicefondo = 1
						multiplicador = -1
						contador = 0
						
						do
							indicefondo = multiplicador + indiceFondo 
							multiplicador = -1 * multiplicador
								
							%>
								<tr class="fondo<%= indicefondo %>">
									<form name="Activar<%= t("CodigoUnico") %>" id="Activar<%= t("CodigoUnico") %>"  method="post" action="activar_codigo.asp">
										<td style="text-align: center; padding: 10px;"><%= t("CodigoUnico") %></td>
										<td style="text-align: left; padding: 10px;"><%= UCase(t("Nombre")) %></td>
										<td style="text-align: center; padding: 10px;"><%= t("Cedula") %></td>
										<td style="text-align: center; padding: 10px;"><%= t("Telefono") %></td>
										<td style="text-align: center; padding: 10px;"><%= UCase(t("Departamento")) %></td>
										
										<td  style="text-align: center; padding: 10px;">
											<img src="imagenes/banderas/<%= t("Bandera") %>" style="width=:33px; height: 22px;"><br />
											<span style="font-size:9px;"><%= t("NombreEquipo") %></span>
										</td>								
										
										<td style="text-align: left; padding: 10px;"><%= t("FechaConfeccion") %></td>
										<td style="text-align: center; padding: 10px;">
											<input name="txtCodigoUnico" type="hidden" value="<%= t("CodigoUnico") %>">
											<% 
												if ver = 0 then 
											%>
													<img src="Imagenes/ver.png"  style="" onClick="VerPolla('<%= t("CodigoUnico") %>')" title="Ver Polla <%= t("CodigoUnico") %>">									
											<%
													if SePuedeActivar() = 1 then
													%>
														<img src='Imagenes/aprobar.png'  style="" onClick="Enviar('Activar<%= t("CodigoUnico") %>')" title='Aprobar Polla <%= t("CodigoUnico") %>'>
													<%
													end if
											%>
													<img src="Imagenes/borrar.png"  style="" onClick="BorrarPolla('<%= t("CodigoUnico") %>')" title="Borrar Polla <%= t("CodigoUnico") %>">	
											<%
												else
											%>
												<img src="Imagenes/ver.png"  style="" onClick="VerPollaBoleto('<%= t("CodigoUnico") %>')" title="Ver Polla <%= t("CodigoUnico") %>">																		
												<img src="Imagenes/boleto.png"  style="" onClick="Boleto('<%= t("CodigoUnico") %>')" title="Ver Boleto <%= t("CodigoUnico") %>">
											<%
												end if
											%>
										</td>								
									</form>
								</tr>
							<%

							contador = contador + 1
							t.movenext
						loop until t.eof
					end if
						
					t.close: set t=nothing
					con.close: set con = nothing
				%>
				
				<tr style="background-color:rgb(0, 0, 0); color:rgb(255, 255, 255);">
					<td colspan="8" style="text-align: center; padding: 10px; font-size: 16px;">
						El Total de Registros es <%= contador %>
					</td>
				</tr>
			</table>

			<br />

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

			function VerPollaBoleto(Codigo) {
				window.location = "boleto.asp?t=" + Codigo
			}

			function Boleto(Codigo) {
				top.location = "boleto.asp?t=" + Codigo
			}			
			
			function BorrarPolla(Codigo) {
				if (confirm('Deseas borrar este registro de la base de datos?')) {
					window.location = "borrar.asp?t=" + Codigo
				} else {
					// No hacer nada...
				}
			}
			
			function Enviar(NombreFormulario) {
				var formulario = document.getElementById(NombreFormulario);
				formulario.submit();
			}
		</script>	
		<!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>	
</html>