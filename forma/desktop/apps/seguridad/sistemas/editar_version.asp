<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Editar Version</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            dim con, t, p, ptt, sqlString, Sistema, Version, ordenadoPor
            dim nombre, Descripcion, ClaseApp, IndiceOrdenamiento, Icono, sBitacora
       
            thisSystem = "seguridad"
            thisProcess = "seg.0090"
            SysLockOut

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")            
        %>
        
        <style>
            .tab {
                padding: 10px; 
                border: 1px solid rgb(187, 188, 189); 
                text-align: center; 
                background-color: rgb(200, 202, 204);                
            }

			.tabDetalles {
				font-family: "Arial Narrow", Arial, sans-serif;
				font-size: 13px; 
				vertical-align: top; 
				padding: 5px; 
				border: 1px solid rgb(187, 188, 189); 
				text-align: left; 
				background-color: rgb(255, 255, 255);
				line-height:2em;
			}   
            
            .borde {
                border: 1px solid;
                border-color: rgb(184, 184, 184);
            }  
            
            .vbControl_B_Enabled {
                background-color: rgb(240, 255, 227);
                color: rgb(0, 0, 0);
                padding: 5px;
                border: 1px solid rgb(28, 69, 117);
            }    
        </style>

        <%
            sub lista_Detalles(Sistema, Version, Activa)
                %>
                    <table style="width: 100%;" style="padding: 0px;"> 
                        <tr style="background-color: rgb(0, 0, 0); color: rgb(255, 255, 255); font-size: 14px;">
                            <td class="top borde" style="width:  5%; text-align: center; padding: 5px;">#</td>
                            <td class="top borde" style="width: 55%; text-align: left;   padding: 5px;">Mejora o Cambio</td>
                            <td class="top borde" style="width: 25%; text-align: left;   padding: 5px;">Solicitado Por</td>
                            <td class="top borde" style="width: 10%; text-align: center; padding: 5px;">Fecha Sol.</td>
                            <td class="top borde" style="width:  5%; text-align: center; padding: 5px;">&nbsp;</td>              
                        </tr>

                        <% if Activa = 0 then %>
                            <tr class="dets">
                                <td class="borde" style="width: 5%; text-align: center; font-size: 30px; font-weight: bold;">
                                    +
                                </td>           

                                <td style='border: solid 1px rgb(180, 180, 180); width: 55%;'>
                                    <input class="field full" id="Descripcion" name="Descripcion" type="text" value= "" required />
                                </td>

                                <td style='border: solid 1px rgb(180, 180, 180); width: 25%;'>
                                    <input class="field full" id="SolicitadoPor" name="SolicitadoPor" type="text" value= "" />
                                </td>

                                <td style='border: solid 1px rgb(180, 180, 180); width: 10%;'>
                                    <input class="field full" style="text-align: center;" id="FechaSolicitado" name="FechaSolicitado" type="text" value= "" placeholder="dd/mm/aaaa" >
                                </td>

                                <td style='border: solid 1px rgb(180, 180, 180); width: 5%;'>
                                    <button type="button" class="form-btn verde" onclick="NuevaLinea('<%= Sistema %>', '<%= Version %>')">
                                        <i class="fa fa-save fa-xl"></i>
                                    </button>
                                </td>
                            </tr>    
                        <% end if %>

                        <%
                            sqlString = "SELECT Version, Sistema, Caracteristica, Descripcion, SolicitadoPor, FechaSolicitado " &  _
                                        "FROM seg_VersionesDetalles " & _
                                        "WHERE (Sistema = '" & Sistema & "') " & _
                                        "AND (Version = '" & Version & "') " & _
                                    "ORDER BY Caracteristica;"

                            set cbox = con.execute(sqlString)
                                if not (cbox.bof or cbox.eof) then
                                    titulo = 0

                                    Do        
                                        titulo = titulo + 1

                                        response.write "<tr>"
                                            response.write "<td style='text-align: center; padding: 5px; width: 5%' class='borde'>"
                                                %><input name="FORM_Caracteristica_<%= titulo %>" id="FORM_Caracteristica_<%= titulo %>" type="text" style="width:0%; text-align: center; visibility: hidden;" value="<%= cbox("Caracteristica") %>"><%
                                                response.write cbox("Caracteristica") 
                                            response.write "</td>"      

                                            response.write "<td style='text-align: left; padding: 5px; width: 48%' class='borde'>"%>
                                                <input name="FORM_Descripcion_<%= titulo %>" id="FORM_Descripcion_<%= titulo %>" type="text" style="width:100%;" value="<%= cbox("Descripcion") %>"><%
                                            response.write "</td>"                                   

                                            response.write "<td style='text-align: left; padding: 5px; width: 25%' class='borde'>"%>
                                                <input name="FORM_SolicitadoPor_<%= titulo %>" id="FORM_SolicitadoPor_<%= titulo %>" type="text" style="width:100%;" value="<%= cbox("SolicitadoPor") %>"><%
                                            response.write "</td>"    

                                            response.write "<td style='text-align: center; padding: 5px; width: 18%' class='borde'>"%>
                                                <input name="FORM_FechaSolicitado_<%= titulo %>" id="FORM_FechaSolicitado_<%= titulo %>" type="text" style="width:100%; text-align: center;" value="<%= SmallFechaFormulario(cbox("FechaSolicitado")) %>"><%
                                            response.write "</td>"                                                                                            

                                            response.write "<td style='text-align: center; padding: 5px; width: 8%' class='borde'>"
                                                %><button class="form-btn rojo" type="button" onClick="BorrarDetalle('<%= Sistema %>', '<%= Version %>', '<%= cbox("Caracteristica") %>')" <% if Activa > 0 then response.write "disabled"%>><%
                                                    response.write "<i class=' fa fa-trash fa-xl' title='Borrar Detalle'></i>"
                                                response.write "</button>"                    
                                            response.write "</td>"                
                                        response.write "</tr>"

                                        cbox.MoveNext
                                    Loop Until (cbox.eof)
                                end if
                            cbox.close: set cbox = nothing     
                        %>                   
                    </table> 
                <%     
            end sub   

            sub Historial(Sistema, Version)
                %>
                    <table style="width: 100%;" style="padding: 0px;"> 
                        <tr style="background-color: rgb(0, 0, 0); color: rgb(255, 255, 255);">
                            <td class="top borde" style="width: 30%; text-align: center; padding: 5px;">Sistema</td>
                            <td class="top borde" style="width: 10%; text-align: center; padding: 5px;">Version</td>
                            <td class="top borde" style="width: 30%; text-align: center; padding: 5px;">Usuario</td>
                            <td class="top borde" style="width: 15%; text-align: center;   padding: 5px;">Notificado</td>
                            <td class="top borde" style="width: 15%; text-align: center; padding: 5px;">Actualizado</td>              
                        </tr>

                        <%
                            sqlString = "SELECT codSistema, Sistema, Version, Usuario, Notificado, Actualizado " & _
                                          "FROM dbo.seg_Versiones_Historial " & _
                                         "WHERE (codSistema = '" & Sistema & "') " & _
                                           "AND (Version = '" & Version & "') " & _
                                      "ORDER BY Usuario;"

                            set cbox = con.execute(sqlString)
                                if not (cbox.bof or cbox.eof) then
                                    titulo = 0

                                    Do        
                                        %>
                                            <tr>
                                            <td class="borde" style="width: 30%; text-align: center; padding: 5px;"><%= cbox("Sistema") %></td>
                                            <td class="borde" style="width: 10%; text-align: center; padding: 5px;"><%= cbox("Version") %></td>
                                            <td class="borde" style="width: 30%; text-align: center; padding: 5px;"><%= cbox("Usuario") %></td>
                                            <td class="borde" style="width: 15%; text-align: center; padding: 5px;"><%= cbox("Notificado") %></td>
                                            <td class="borde" style="width: 15%; text-align: center; padding: 5px;"><%= cbox("Actualizado") %></td>
                                            </tr>
                                        <%

                                        cbox.MoveNext
                                    Loop Until (cbox.eof)
                                end if
                            cbox.close: set cbox = nothing     
                        %>                   
                    </table> 
                <%   
            end sub            

            function NombreSistema(Sistema)
                dim sqlString

                sqlString = "SELECT sysNombre FROM seg_Sistemas WHERE (sysCodigo ='" & Sistema & "');"

                set ptt = con.execute(sqlString)   
                    NombreSistema = ptt("sysNombre") 
                ptt.close: set ptt = nothing
            end function     

            function FechaFormulario(FechaSQL)       
                dim d, m, a, h, mm

                if FechaSQL <> "" then
                    a = YEAR(FechaSQL)
                    m = RIGHT("00" & MONTH(FechaSQL), 2)
                    d = RIGHT("00" & DAY(FechaSQL), 2)

                    h = RIGHT("00" & HOUR(FechaSQL), 2)
                    mm = RIGHT("00" & MINUTE(FechaSQL), 2)

                    FechaFormulario = d & "/" & m & "/" & a & " " & h & ":" & mm
                else
                    FechaFormulario = ""
                end if
            end function    

            function SmallFechaFormulario(FechaSQL)       
                dim d, m, a

                if FechaSQL <> "" then
                    a = YEAR(FechaSQL)
                    m = RIGHT("00" & MONTH(FechaSQL), 2)
                    d = RIGHT("00" & DAY(FechaSQL), 2)

                    SmallFechaFormulario = d & "/" & m & "/" & a 
                else
                    SmallFechaFormulario = ""
                end if
            end function                                 
        %>
    </head>

    <body plantilla="normal" reserva="175" onload="init()">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->   

        <%
            Sistema = request.querystring("s")
            Version = request.querystring("v")
            ordenadoPor = request.querystring("o")
        %>

        <br />

        <div style="width: 95%; margin: auto;">
            <!--
                El valor de ACTIVA modifica el formulario:

                0: La versión está en proceso y se puede editar libremente.
                   No se reporta a los usuarios y ellos no pueden verla.

                1: Es la "Versión Actual". Es la que se usa para comprobar
                   la versión que usa el usuario. Ya no se puede editar y es
                   la que ve el usuario al hacer login.

                2: Es una versión que fue desplazada por otra. Aparece en el
                   historial de versiones y el usaurio puede consultarla.
            -->

            <%
                sqlString = "SELECT Resumen, Obligatoria, Activa, FechaActivacion " & _
                            "FROM seg_Versiones " & _
                            "WHERE (Sistema = '" & Sistema & "') " & _
                            "AND (Version = '" & Version & "');"
                
                set t = con.execute(sqlString)                
            %>
      
            <table style="width: 95%;">
                <tr>
                    <td style="width: 30%; font-size: 24px;">
                        <% response.write "Editar Version " & Version & "<br/>(" & NombreSistema(Sistema) & ")" %>
                    </td>

                    <td style="width: 70%; text-align: right;">
                        <button class='form-btn rojo normal'  type='button' onclick="volver()">Cancelar</button>&nbsp;&nbsp;
                        <% if t("Activa") = 0 then %>
                            <button class='form-btn verde normal' type='button' onclick="grabar()">Grabar</button>&nbsp;&nbsp;
                        <% end if %>                            
                    </td>
                </tr>
            </table>    


            <form id="formulario"  name="formulario" method="post" action="grabar_version.asp">
                <input id="ordenadoPor" name="ordenadoPor" type="text" value="<%= ordenadoPor %>" class="no-ver"/>
                <input id="Sistema" name="Sistema" type="text" value="<%= Sistema %>" class="no-ver"/>
                <input id="Version" name="Version" type="text" value="<%= Version %>" class="no-ver"/>

                <div class="main main-scroll">
                    <div class="line">
                        <label class="label large">Version</label>
                        <input class="field normal" id="Version_display" name="Version_display" value = "<%= Version %>" type="text" disabled />
                    </div>

                    <div class="line">
                        <label class="label large">Resumen</label>
                        <input class="field xxl" id="Resumen" name="Resumen" type="text" value = "<%= t("Resumen") %>" required <%
                            if t("Activa") > 0  then
                                response.write "disabled"
                            end if
                        %>/>
                    </div>

                    <div class="line">
                        <label class="label large">Obligatoria</label>
                        <select class="field normal" name="Obligatoria" id="Obligatoria" <%
                            if t("Activa") > 0 then
                                response.write "disabled"
                            end if
                        %>>
                            <option value="0" <% if t("Obligatoria") = 0 then response.write "selected" %>>&nbsp;</option>
                            <option value="1" <% if t("Obligatoria") = 1 then response.write "selected" %>>Es Obligatoria</option>
                        </select>
                    </div>

                    <% if t("Activa") > 0 then %>
                        <div class="line">
                            <label class="label large">Fecha de Activacion</label>
                            
                            <input class="field small" id="Fecha" name="Fecha" type="text" value = "<%= FechaFormulario(t("FechaActivacion")) %>" <%
                                if t("Activa") > 0  then
                                    response.write "disabled"
                                end if
                            %>/>                                                                                 
                        </div>                                    
                    <% end if %>


                    <!-- Inicio de Tabs -->

                        <div class="line">
                            <label class="label full section"> 
                                <table style="width: 100%; border: none; border-spacing: 0px;">
                                    <tbody>
                                        <tr>
                                            <td style="padding: 10px; border: 1px solid rgb(187, 188, 189); text-align: center; background-color: rgb(200, 202, 204);"
                                                onclick="Tabs_Display('tab_detalles')">
                                                Detalles
                                            </td>

                                            <td style="padding: 10px; border: 1px solid rgb(187, 188, 189); text-align: center; background-color: rgb(200, 202, 204);"
                                                onclick="Tabs_Display('tab_historial')">
                                                Historial
                                            </td>
                                        </tr>

                                        <tr>
                                            <td colspan="2" style="padding: 10px; border: 1px solid rgb(187, 188, 189); text-align: center; background-color: rgb(255, 255, 255);">
                                                <div id="tab_detalles" style="display: none; text-align: left; font-size: 16px; line-height: 1.8em;">
                                                    <% lista_Detalles Sistema, Version, t("Activa") %>
                                                </div>

                                                <div id="tab_historial" style="display: none; text-align: left; font-size: 16px; line-height: 1.8em;">
                                                    <% Historial Sistema, Version %>
                                                </div>
                                            </td>
                                        </tr>
                                    </tbody>
                                </table>
                            </label>
                        </div>

                    <!-- Fin de Tabs -->
                </div>
            </form>
        </div>

        <script type="text/javascript">
            pageReserva = 165;

            function init() {
                Tabs_Display('tab_detalles');
            }

            function volver() {
                var vinculo = "versiones.asp?s=<%= Sistema %>&o=<%= ordenadoPor %>";
                window.location.href = vinculo;
            }

            function grabar() {
                document.getElementById("formulario").submit();
            }    
            
            function NuevaLinea(sistema, version) {
                var descripcion = document.getElementById("Descripcion").value;
                var solicitadoPor = document.getElementById("SolicitadoPor").value;
                var fechaSolicitado = document.getElementById("FechaSolicitado").value;
                
                descripcion = descripcion.replace(/&/g, "yy");

                if (descripcion != "") {
                    var vinculo = "editar_version_detalle.asp?s=" + sistema + "&v=" + version + "&d=" + descripcion +"&sp=" + solicitadoPor + "&fs=" + fechaSolicitado;
                    window.location.href = vinculo;
                };               
            }     
            
            function BorrarDetalle(sistema, version, caracteristica) {
                var confirmacion = confirm("Desea borrar el detalle seleccionado?");
                var vinculo ="borrar_version_detalle.asp?s=" + sistema + "&v=" + version + "&c=" + caracteristica;

                if (confirmacion) {     
                    window.location.href = vinculo;
                };
            }               

            function Tabs_Display(codigo) {
                var h = document.getElementById("tab_historial");
                var d = document.getElementById("tab_detalles");

                switch (codigo) {
                    case "tab_detalles":
                        h.style.display = 'none';
                        d.style.display = 'block';
                        break;
                    case "tab_historial":
                        d.style.display = 'none';
                        h.style.display = 'block';                        
                        break;
                }
            }

            mask(document.getElementById('FechaSolicitado'), ['99/99/9999']);
        </script> 

        <%
            t.close: set t = nothing
            con.close: set con = nothing
        %>          
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->    	        
    </body>
</html>