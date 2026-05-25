<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
  	<head>
		<meta charset="utf-8">

		<title>Editar Roles</title>
		<!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0080"
            SysLockOut

			dim pCon, ptt, pProc, pTemp, sqlString, Rol, ordenadoPor
			dim Nombre, Descripcion, TipoRol, sysCodigo

			Rol = request.querystring("r")
			ordenadoPor = request.querystring("o")
			sysCodigo = request.querystring("s")

			set pCon = Server.CreateObject("ADODB.Connection")
			pCon.open Application("Conn")

			function NombreSistema(Sistema)
				dim sqlString

				sqlString = "SELECT sysNombre FROM seg_Sistemas WHERE (sysCodigo ='" & Sistema & "');"

				set pTemp = pCon.execute(sqlString)   
					NombreSistema = pTemp("sysNombre") 
				pTemp.close: set pTemp = nothing
			end function     

			function CuantosProcesos(Sistema)
				dim sqlString

				sqlString = "SELECT Count(*) AS Cuantos FROM seg_Procesos WHERE (proSistema ='" & Sistema & "');"

				set pTemp = pCon.execute(sqlString)   
					CuantosProcesos = pTemp("Cuantos") 
				pTemp.close: set pTemp = nothing
			end function    			
		%>

		<style>
			.tabDetalles {
				font-family: "Arial Narrow", Arial, sans-serif;
				font-size: 13px; 
				vertical-align: top; 
				/* border: 1px solid rgb(212, 212, 212); */
				padding: 5px; 
				text-align: left; 
				background-color: rgb(255, 255, 255);
				line-height:2em;
			}
        </style>
  	</head>

  	<body plantilla="normal" reserva="165">
		<!-- #include virtual = "/core/includes/kernel/body.inc" -->   

		<%
			if Rol <> "*" then
				set ptt = pCon.execute("SELECT * FROM seg_Roles WHERE (rolCodigo = '" & Rol & "');")
					Nombre = ptt("rolNombre")
					Descripcion = ptt("rolDescripcion")
					TipoRol = ptt("TipoRol")
				ptt.close: set ptt = nothing
			end if
		%>

		<div style="width: 95%; margin: auto;">
			<br />

			<table style="width: 95%"> 
				<tr>
					<td style="width: 30%; font-size: 24px;">
						<%
							if Rol = "*" then
								response.write "Nuevo Rol"
							else
								response.write Nombre
							end if
						%>
					</td>

					<td style="width: 70%; text-align: right;">
						<button class='form-btn rojo normal'  type='button' onclick="volver()">Cancelar</button>&nbsp;&nbsp;
						<button class='form-btn verde normal' type='button' onclick="grabar()">Grabar</button>&nbsp;&nbsp;
					</td>
				</tr>
			</table>   
		</div>

		<div style="width: 98%; margin: auto;">
			<form id="formulario"  name="formulario" method="post" action="grabar_rol.asp">
				<input id="ordenadoPor" name="ordenadoPor" type="text" value="<%= ordenadoPor %>" class="no-ver"/>
				<input id='unico' name='unico' type='text' value='1' class="no-ver"/>

				<div class="main main-scroll">
					<div class="line">			
						<% if Rol = "*" then %>
							<input id='nuevo' name='nuevo' type='text' value='1' class="no-ver"/>
						
							<label class="label normal">Codigo</label>
							<input class="field normal" id="Codigo" name="Codigo" type="text" required />
						<% else %>
							<input id='nuevo' name='nuevo' type='text' value='0' class="no-ver"/>
							<input id="Codigo" name="Codigo" type="text" value="<%= Rol %>" class="no-ver"/>

							<label class="label normal">Codigo</label>
							<input class="field normal" id="dispCodigo" name="dispCodigo" type="text" value='<%= Rol %>' disabled />
						<% end if %>
					</div>

				<div class="line">
					<label class="label normal">Nombre</label>
					<input class="field large" id="Nombre" name="Nombre" type="text" required <% if Rol <> "*" then response.write "value='" & Nombre & "'" %> />
				</div>

				<div class="line">
					<label class="label normal">Descripcion</label>
					<input class="field xxl" id="Descripcion" name="Descripcion" type="text" required <% if Rol <> "*" then response.write "value='" & Descripcion & "'" %> />
				</div>                      

				<div class="line">
					<label class="label normal">Tipo de Rol</label>
					<select class="field normal" name="TipoRol" id="TipoRol">
						<option value="1" <% if Rol <> "*" then 
												if TipoRol = 1 then response.write "selected"
											end if
											%> >Rol</option>

						<option value="0" <% if Rol <> "*" then 
												if TipoRol = 0 then response.write "selected"
											else
												response.write "selected"
											end if
											%> >Anti-Rol</option>
					</select>
				</div>


				<!--
				'
				' PROCESOS DE LOS SISTEMAS
				'
				-->

				<div class="line label-top">
					<label class="label normal">Procesos</label>
					<div class="label full section"> 
						<%
							response.write "<table style='width: 100%; border: none; border-spacing: 0px; text-align: center;'>"
								response.write "<tr>"

								if CuantosProcesos(sysCodigo) > 5 then
									procs = CuantosProcesos(sysCodigo)
									cols = int(procs / 5)

									cols = cols + (-1 * (procs > cols))
								else
									cols = 1
								end if
									
								sqlString = "SELECT p.proSistema, p.proCodigo, p.proNombre AS NomProceso, " & _
												" CASE WHEN detRol IS NULL THEN 0 ELSE 1 END AS Selected " & _
											"FROM dbo.seg_Procesos AS p " & _
											"LEFT OUTER JOIN (SELECT detRol, detRolSistema, detRolProceso " & _
															"FROM dbo.seg_RolDetalles " & _
															"WHERE (detRol = '" & Rol & "')) AS r " & _
											"ON p.proSistema = r.detRolSistema " & _
											"AND p.proCodigo = r.detRolProceso " & _
											"WHERE (p.proMenuItem = 1) " & _
											"AND (p.proSistema = '" & sysCodigo & "') " & _
											"ORDER BY NomProceso;"

								set pProc = pCon.execute(sqlString)
									if not (pProc.bof or pProc.eof) then
										contador = 0
										td = 1

										response.write "<td class = 'tabDetalles'>"
											do
												contador = contador + 1
												if contador > 5 then
													contador = 1
													response.write "</td>"
													response.write "<td class = 'tabDetalles'>"
												end if

												pProc_id = pProc("proSistema") & "__" & pProc("proCodigo") 
												pProc_id = Replace(pProc_id, ".", "_")
													%>
														<input type='checkbox' id='<%= pProc_id %>' name='<%= pProc_id %>' value='1' 
															<% if pProc("Selected") = 1 then response.write "checked" %>>
														<%= pProc("NomProceso") %>
													<%
												response.write "<br/>"

												pProc.movenext
											loop until pProc.eof
									end if
								pProc.close: set pProc = nothing  


									response.write "</td>"
								response.write "</tr>"
							response.write "</table>"
						%>                                          
					</div>
				</div>

				<!--
				'
				' FIN DE PROCESOS
				'
				-->
			</form>
		</div>

		<br /><br />

		<script>
			function volver() {
				var vinculo = "roles_sys.asp?s=<%= sysCodigo %>&o=<%= ordenadoPor %>";
				window.location.href = vinculo;
			}

			function grabar() {
				document.getElementById("formulario").submit();
			}        
		</script> 

        <% pCon.close: set pCon = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->    		
  	</body>
</html>