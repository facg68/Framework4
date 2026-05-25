<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
  	<head>
		<meta charset="utf-8">

		<title>Editar Variables</title>
		<!-- #include virtual = "/core/includes/kernel/head.inc" -->  

        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0085"
            SysLockOut
			
			dim pCon, ptt, pProc, pTemp, sqlString, OrdenadoPor, lista
			dim Sistema, Variable, ValorDefault, ConsultaSQL, Descripcion, Afectacion, Tipo

			Sistema = request.querystring("s")
			Variable = request.querystring("p")
			ordenadoPor = request.querystring("o")

			set pCon = Server.CreateObject("ADODB.Connection")
			pCon.open Application("Conn")
		%>		
  	</head>

  	<body plantilla="normal" reserva="165">
		<!-- #include virtual = "/core/includes/kernel/body.inc" -->   

		<%
			if Variable <> "*" then
				set ptt = pCon.execute("SELECT * FROM seg_Parametros WHERE (Sistema = '" & Sistema & "') AND (Parametro = '" & Variable & "');")
					Tipo = ptt("TipoParametro")
					ValorDefault = ptt("ValorDefault")
					ConsultaSQL = ptt("ConsultaSQL")
					Exponer = ptt("Exponer")
					Descripcion = ptt("Descripcion")
					Afectacion = ptt("Afectacion")
				ptt.close: set ptt = nothing
			else	
				Tipo = 1
			end if
		%>

		<div style="width: 95%; margin: auto;">
			<br />

			<table style="width: 100%"> 
				<tr>
					<td style="width: 30%; font-size: 24px;">
						<%
							if Variable = "*" then
								response.write "Nueva Variable"
							else
								response.write Variable
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

		<div style="width: 95%; margin: auto;">
			<form id="formulario"  name="formulario" method="post" action="grabar_variable.asp">
				<input id="ordenadoPor" name="ordenadoPor" type="text" value="<%= ordenadoPor %>" class="no-ver"/>

				<div class="main main-scroll">
					<div class="line">
						<% if Variable = "*" then %>
								<input id='nuevo' name='nuevo' type='text' value='1' class="no-ver"/>
								<input id='sistema' name='sistema' type='text' value='<%= sistema %>' class="no-ver"/>
								
								<label class="label normal">Variable</label>
								<input class="field normal" id="Parametro" name="Parametro" type="text" required />
						<% else %>
								<input id='nuevo' name='nuevo' type='text' value='0' class="no-ver"/>
								<input id='sistema' name='sistema' type='text' value='<%= sistema %>' class="no-ver"/>
								<input id="Parametro" name="Parametro" type="text" value="<%= Variable %>" class="no-ver"/>

								<label class="label normal">Codigo</label>
								<input class="field normal" id="dispCodigo" name="dispCodigo" type="text" value='<%= Variable %>' disabled />
						<% end if %>
					</div>

					<div class="line">
						<label class="label normal">Descripcion</label>
						<input class="field xl" id="Descripcion" name="Descripcion" type="text" value='<%= Descripcion %>' />
					</div>    											

					<div class="line">
						<label class="label normal">Tipo de Variable</label>
						
						<select class="field normal" name='TipoParametro' id='TipoParametro' onChange="VerificarLista()">
							<option value='1' <% if Tipo = "1" then response.write " selected" %>>Permiso o Restricción</option>
							<option value='2' <% if Tipo = "2" then response.write " selected" %>>Variable</option>
							<option value='3' <% if Tipo = "3" then response.write " selected" %>>Sí / No</option>
							<option value='4' <% if Tipo = "4" then response.write " selected" %>>Lista</option>
							<option value='5' <% if Tipo = "5" then response.write " selected" %>>Consulta SQL</option>
							<option value='6' <% if Tipo = "6" then response.write " selected" %>>Selección de Color</option>
							<option value='7' <% if Tipo = "7" then response.write " selected" %>>Barra de Desplazamiento</option>
						<select>
					</div>     

					<div id="div_lista" style="display: none;">
						<div class="line">
							<label class="label normal">Valores</label>
							<div class="label full section"> 
								<table style="width: 100%; style='border: solid 1px rgb(180, 180, 180);'">
									<tr style="background-color: rgb(180, 180, 180); color: black; font-size: 14px; text-align: center;">
										<td style='width:30%;'>Valor</td>
										<td style='width:60%;'>Descripción</td>
										<td style='width:10%;'>Def</td>
									</tr>
									<%
										if Variable = "*" then
											for k = 1 to 15
												response.write "<tr>"
													response.write "<td style='border: solid 1px rgb(180, 180, 180);'><input class='field' style='width: 100%;' 	id='lValor_" & k & "' name='lValor_" & k & "' type='text' /></td>"
													response.write "<td style='border: solid 1px rgb(180, 180, 180);'><input class='field' style='width: 100%;' 	id='lDescripcion_" & k & "' name='lDescripcion_" & k & "' type='text' /></td>"
													response.write "<td style='border: solid 1px rgb(180, 180, 180); text-align: center;'><input type='checkbox' 	id='lPorDefecto_" & k & "' name='lPorDefecto_" & k & "' value='1' onClick='resetChk(" & k & ")' /></td>"
												response.write "</tr>"
											next
										else
											set ll = pCon.execute("SELECT * FROM seg_Parametros_Valores WHERE Sistema = '" & Sistema & "' AND Parametro = '" & Variable & "' ORDER BY Valor;")
												if not (ll.bof or ll.eof) then
													k = 0
													do
														k = k + 1

														response.write "<tr>"
															response.write "<td style='border: solid 1px rgb(180, 180, 180);'>" 
																response.write "<input class='field full' id='lValor_" & k & "' name='lValor_" & k & "' type='text' value= '" & ll("Valor") & "' />"
															response.write "</td>"

															response.write "<td style='border: solid 1px rgb(180, 180, 180);'>"
																response.write "<input class='field full' id='lDescripcion_" & k & "' name='lDescripcion_" & k & "' type='text' value= '" & ll("Descripcion") & "'/>"
															response.write "</td>"

															response.write "<td style='border: solid 1px rgb(180, 180, 180); text-align: center;'>" 
																response.write "<input type='checkbox' id='lPorDefecto_" & k & "' name='lPorDefecto_" & k & "' value='1' "
																	if ll("porDefecto") = 1 then
																		response.write " checked "
																	end if
																response.write " onClick='resetChk(" & k & ")' />"
															response.write "</td>"
														response.write "</tr>"

														ll.MoveNext
													Loop until ll.eof
												end if
											ll.close: set ll = nothing

											if k <> 15 then
												for k1 = (k + 1) to 15
													response.write "<tr>"
														response.write "<td style='border: solid 1px rgb(180, 180, 180);'><input class='field full' id='lValor_" & k1 & "' name='lValor_" & k1 & "' type='text' /></td>"
														response.write "<td style='border: solid 1px rgb(180, 180, 180);'><input class='field full' id='lDescripcion_" & k1 & "' name='lDescripcion_" & k1 & "' type='text' /></td>"
														response.write "<td style='border: solid 1px rgb(180, 180, 180); text-align: center;'><input type='checkbox' id='lPorDefecto_" & k1 & "' name='lPorDefecto_" & k1 & "' value='1' onClick='resetChk(" & k1 & ")' /></td>"
													response.write "</tr>"																	
												next
											end if
										end if
									%>
								</table>
							</div>
						</div>
					</div>    

					<div id="div_consulta" style="display: none;">
						<div class="line">
							<label class="label normal">
								Consulta SQL
								<br /><br />
								El resultado siempre debe ser sólo dos columnas: [Codigo] y [Valor]
							</label>
							
							<label class="label full section"> 
								<textarea style="border: solid 1px rgb(210, 210, 210);font-family: 'Courier New'; font-size: 14px; width: 100%;" 
											name="ValorConsulta" 
											id ="ValorConsulta" 
											rows="15" cols="80"><% if Variable <> "*" and Tipo = "5" then response.write ConsultaSQL %></textarea>
							</label>
						</div>
					</div>																																	

					<div id="div_default" style="display: none;">
						<div class="line">
							<label class="label normal">Valor Default</label>
							<input class="field normal" id="ValorDefault" name="ValorDefault" type="text" required 
								<% 
									if Variable <> "*" and Tipo = "2" then 
										response.write "value='" & ValorDefault & "'" 
									end if
								%> 
							/>
						</div>
					</div>

					<div id="div_colores" style="display: none;">
						<div class="line">
							<label class="label normal">Color Default</label>
							<%
								divCol = "#ffffff"

								if (Variable <> "*") AND (len(trim(ValorDefault)) > 0) then
									divCol = ValorDefault 
								end if														
							%>
							<input class="field normal" type="color" 
									id="ColorPredeterminado" name="ColorPredeterminado" 
									value="<%= divCol %>"
									style="background-color: <%= divCol %> ;"
									onChange = "reColor('ColorPredeterminado')"
							>
						</div>
					</div>	

					<div id="div_desplazamiento" style="display: none;">
						<div class="line">
							<label class="label normal">Valor</label>
							
							<label class="label full section"> 
								<div style="width: 90%; text-align: left;">
									<input class="field normal" style="width:95%" type="range" min="0" max="100" 
										value="<%
											if Variable = "*" then
												response.write "30"
											else
												response.write ValorDefault 
											end if														   	
										%>" 
										id="barraDesplazamiento" name="barraDesplazamiento">
								</div>
								
								<label style="width:5%" id="BarraValor"></label>  

								<script>
									var slider = document.getElementById("barraDesplazamiento");
									var output = document.getElementById("BarraValor");

									output.innerHTML = slider.value;

									slider.oninput = function() {
										output.innerHTML = this.value;
									}
								</script>   
							</label>
						</div>
					</div>																						

					<div class="line">
						<label class="label normal">Exponer</label>
						
						<select class="field large" name='Exponer' id='Exponer'>
							<option value='0' <% if Exponer = "0" then response.write " selected" %>>&nbsp;</option>
							<option value='1' <% if Exponer = "1" then response.write " selected" %>>El Usuario Puede Modificarla</option>
						<select>
					</div>   

					<div class="line">
						<label class="label normal">Cómo Afecta</label>
						<div class="label full section"> 
							<script src="/core/lib/tinymce/tinymce.min.js"></script>                        
							
							<textarea class="editor" id="Afectacion" name="Afectacion"> 
								<%= Afectacion %>
							</textarea>
							<script>
								tinymce.init({
								entity_encoding : "raw",
								selector: '.editor',
								license_key: 'gpl',
								height: 400,
								branding: false,
								promotion: false,								
								language: 'es',
								language_url: '/core/includes/es.js', 
								plugins: 'anchor autolink charmap codesample emoticons image link lists media searchreplace table visualblocks wordcount ',
								toolbar: 'undo redo | blocks fontfamily fontsize | bold italic underline strikethrough | link image media table mergetags | addcomment showcomments | spellcheckdialog a11ycheck typography | align lineheight | checklist numlist bullist indent outdent | emoticons charmap | removeformat'
								});
							</script>   
						</div>
					</div>  
				</div>  
			</form>
		</div>

		<br /><br />

		<script>
			function VerificarLista() {
				var valor = document.getElementById("TipoParametro").value;
				document.getElementById("div_default").style.display = "none";
				document.getElementById("div_lista").style.display = "none"; 
				document.getElementById("div_consulta").style.display = "none"; 
				document.getElementById("div_colores").style.display = "none"; 
				document.getElementById("div_desplazamiento").style.display = "none"; 

				if (valor == 2) { document.getElementById("div_default").style.display = "block"; };
				if (valor == 4) { document.getElementById("div_lista").style.display = "block"; };
				if (valor == 5) { document.getElementById("div_consulta").style.display = "block"; };
				if (valor == 6) { document.getElementById("div_colores").style.display = "block"; };
				if (valor == 7) { document.getElementById("div_desplazamiento").style.display = "block"; };
			}

			function resetChk(campo) {
				var k, nombre;

				for (let k = 1; k < 16; k++) {
					nombre = "lPorDefecto_" + k;
					document.getElementById(nombre).checked = false;
				}				

				nombre = "lPorDefecto_" + campo;
				document.getElementById(nombre).checked = true;
			}
		
			function volver() {
				var vinculo = "variables.asp?s=<%= Sistema %>&o=<%= ordenadoPor %>";
				window.location.href = vinculo;
			}

			function grabar() {
				var consulta = document.getElementById("ValorConsulta").value;
				document.getElementById("ValorConsulta").value = getCharCodes(consulta);
				document.getElementById("formulario").submit();
			}      

			function getCharCodes(s){
				let charCodeArr = [];
				
				for(let i = 0; i < s.length; i++){
					let code = s.charCodeAt(i);
					charCodeArr.push(code);
				}
				
				return charCodeArr;
			}

            function reColor(Campo) {
                const elemento = document.getElementById(Campo);
                const color = elemento.value;
                
                if (!/^#[0-9A-Fa-f]{6}$/.test(color)) {
                    console.warn(`El valor "${color}" no es un color hexadecimal válido`);
                    return;
                }

                elemento.style.backgroundColor = color;
            }	
			
			window.addEventListener('load', function() {
				VerificarLista();				
			});			
		</script> 

        <% pCon.close: set pCon = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->    		  
  	</body>
</html>