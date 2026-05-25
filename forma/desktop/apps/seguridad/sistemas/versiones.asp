<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Versiones de los Sistemas</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->            

        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0090"
            SysLockOut

            dim cc, t, tt, sqlString, ordenadoPor, Sistema

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")            

            function SistemaDefault()
                dim sqlString

                sqlString = "SELECT TOP (1) sysCodigo FROM seg_Sistemas ORDER BY sysNombre;"

                set tt = cc.execute(sqlString)   
                    SistemaDefault = tt("sysCodigo") 
                tt.close: set tt = nothing
            end function       

            function NombreSistema(Sistema)
                dim sqlString

                sqlString = "SELECT sysNombre FROM seg_Sistemas WHERE (sysCodigo ='" & Sistema & "');"

                set tt = cc.execute(sqlString)   
                    NombreSistema = tt("sysNombre") 
                tt.close: set tt = nothing
            end function    

            function FechaFormulario(FechaServer)
                dim dd, mm, aa, hh, min

                if FechaServer <> "" then
                    aa = left(FechaServer, 4)
                    mm = right("00" & mid(FechaServer, 6, 2), 2)
                    dd = right("00" & mid(FechaServer, 9, 2), 2)

                    hh = right("00" & mid(FechaServer, 12, 2), 2)
                    min = right("00" & mid(FechaServer, 15, 2), 2)

                    FechaFormulario = dd & "/" & mm & "/" & aa & "<br/>" & hh & ":" & min
                else
                    FechaFormulario = "&nbsp;"
                end if
            end function
        %>    
    </head>

    <body plantilla="lista" reserva="225">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    

        <br />

        <%
            Sistema = request.querystring("s")                    
            ordenadoPor = request.querystring("o")                    

            if Sistema = "" then Sistema = SistemaDefault()
            if ordenadoPor = "" then ordenadoPor = "0"

            sqlString = "SELECT Version, Sistema, Resumen, Obligatoria, Activa, FechaActivacion, dbo.FechaFormulario(FechaActivacion) AS fForm  " & _
                          "FROM seg_versiones " & _
                          "WHERE Sistema = '" & Sistema & "' "

            select case ordenadoPor
                case 0: sqlString = sqlString & " ORDER BY Version;"
                case 1: sqlString = sqlString & " ORDER BY Resumen;"
                case 2: sqlString = sqlString & " ORDER BY FechaActivacion;"                
                case 3: sqlString = sqlString & " ORDER BY Version DESC;"                
                case 4: sqlString = sqlString & " ORDER BY Resumen DESC;"
                case 5: sqlString = sqlString & " ORDER BY FechaActivacion DESC;"                                
            end select

            set t = cc.execute(sqlString)        
        %>

        <div style="width:98%; margin: auto;">
            <table style="width:100%; margin: auto;">
                <tr style="padding: 10px;">
                    <td>
                        <table style="width: 100%;">                
                            <tr>
                                <td style="text-align:left; width: 60%">
                                    <span style="font-size: 24px">
                                        &nbsp;Versiones - <%= NombreSistema(Sistema) %>
                                    </span>                                

                                    <br />

                                    <span style="font-size: 20px">
                                        &nbsp;Ordenadas por&nbsp;<%
                                            select case ordenadoPor
                                                case 0: response.write "Versión"
                                                case 1: response.write "Resumen"
                                                case 2: response.write "Fecha de Activación"
                                                case 3: response.write "Versión (descendentemente)"
                                                case 4: response.write "Resumen (descendentemente)"
                                                case 5: response.write "Fecha de Activación (descendentemente)"
                                            end select                            
                                        %>
                                    </span>                            
                                </td>

                                <td style="text-align:right; width: 40%">
                                    <select name="cboSistema" id="cboSistema" required 
                                            class="field"
                                            onChange="filtrar()">

                                            <%
                                                set tt = cc.execute("SELECT sysCodigo AS Codigo, sysNombre AS Nombre FROM seg_Sistemas ORDER BY Nombre;")
                                                    if not (tt.bof or tt.eof) then
                                                        Do
                                                            response.write "<option value='" & tt("Codigo") & "' "
                                                                if tt("Codigo") = Sistema then 
                                                                    response.write " selected"
                                                                end if
                                                            response.write ">" & tt("Nombre") & "</option>"

                                                            tt.MoveNext
                                                        Loop Until tt.eof
                                                    end if
                                                tt.close: set tt = nothing
                                            %>
                                    </select>  

                                    &nbsp;&nbsp;    

                                    <select name="ordenadoPor" id="ordenadoPor" required 
                                            class="field"
                                            onChange="filtrar()">
                                        <option value="0" <% if ordenadoPor = "0" then response.write " selected" %>>&#9650; Versión</option>
                                        <option value="1" <% if ordenadoPor = "1" then response.write " selected" %>>&#9650; Resumen</option>
                                        <option value="2" <% if ordenadoPor = "2" then response.write " selected" %>>&#9650; Fecha de Activación</option>
                                        <option value="3" <% if ordenadoPor = "3" then response.write " selected" %>>&#9660; Versión</option>
                                        <option value="4" <% if ordenadoPor = "4" then response.write " selected" %>>&#9660; Resumen</option>
                                        <option value="5" <% if ordenadoPor = "5" then response.write " selected" %>>&#9660; Fecha de Activación</option>
                                    </select>         

                                    &nbsp;&nbsp;

                                    <button type="button" class="form-btn verde" onclick="nueva()">
                                        <i class=" fa fa-edit fa-xl" title="Nueva"></i>
                                    </button>                                 
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>

            <table style="width:100%; margin: auto;">                
                <tr style="font-size: 14px; background-color: rgb(61, 61, 61); color:rgb(255,255,255);">
                    <td>
                        <table style="width: 100%;">                
                            <tr style="font-size: 14px; background-color: rgb(61, 61, 61); color:rgb(255,255,255);">
                                <td style="padding: 10px; text-align: center; width: 15%">Versión</td>
                                <td style="padding: 10px; text-align:   left; width: 25%">Resumen</td>
                                <td style="padding: 10px; text-align: center; width: 10%">Obligatoria</td>
                                <td style="padding: 10px; text-align: center; width: 15%">Estado</td>
                                <td style="padding: 10px; text-align: center; width: 20%">F.Act.</td>

                                <td style="padding: 10px; text-align:center; width:  15%">&nbsp;</td>
                            </tr>
                        </table>
                    </td>
                </tr>

                <tr>
                    <td>
                        <div id="overFlow" style="width:100%; height: 650px; overflow: auto; background-color: rgb(207, 207, 207);">                        
                            <table style="width: 100%;">
                                <%  
                                    if not (t.bof or t.eof) then 
                                        conta = 0 
                                        Do     
                                            conta = conta + 1
                                %>
                                        <tr style="font-size: 14px; background-color: rgb(255,255,255); color:rgb(0,0,0); border-bottom: 1px solid rgb(194, 194, 194);">                                
                                            <td style="padding: 5px; text-align:center; width: 15%;" onclick="editar('<%= t("Version") %>')"><%= t("Version") %></td>
                                            <td style="padding: 5px; text-align:center; width: 25%;" onclick="editar('<%= t("Version") %>')"><%= t("Resumen") %></td>
                                            <td style="padding: 5px; text-align:center; width: 10%;" onclick="editar('<%= t("Version") %>')">
                                                <%
                                                    if t("Obligatoria") = 1 then
                                                        response.write "Obligatoria"
                                                    else
                                                        response.write "&nbsp;"
                                                    end if
                                                %>
                                            </td>
                                            <td style="padding: 5px; text-align:center; width: 15%;" onclick="editar('<%= t("Version") %>')">
                                                <%
                                                    select case t("Activa")
                                                        case 0: response.write "En Proceso"
                                                        case 1: response.write "Actual"
                                                        case 2: response.write "Obsoleta"
                                                    end select
                                                %>
                                            </td>
                                            <td style="padding: 5px; text-align:center; width: 20%;" onclick="editar('<%= t("Version") %>')"><%= t("fForm") %></td>

                                            <td style="padding: 5px; text-align: right; width: 15%;">
                                                <button type="button" class="form-btn azul" onclick="activar('<%= t("Version") %>')" <% if t("Activa") > 0 then response.write "disabled" %>>
                                                    <i class=" fa fa-check fa-xl" title="Activar Version"></i>
                                                </button>

                                                <button type="button" class="form-btn rojo" onclick="borrar('<%= t("Version") %>')" <% if t("Activa") > 0 then response.write "disabled" %>>
                                                    <i class=" fa fa-trash fa-xl" title="Borrar Sistema"></i>
                                                </button>                                                
                                            </td>
                                        </tr>
                                <% 
                                            t.MoveNext
                                        Loop Until (t.eof)
                                    end if 
                                %>
                            </table>
                        </div>                
                    </td>
                </tr>

                <tr style="font-size: 14px; background-color: rgb(61, 61, 61); color:rgb(255,255,255);">
                    <td style="padding: 10px; text-align:center; width: 100%;">
                        <% 
                            if conta = 0 then   
                                response.write "No se encontró ninguna versión"
                            else
                                if conta = 1 then
                                    response.write "Se encontró una versión"
                                else
                                    response.write "Se encontraron " & conta & " versiones"
                                end if
                            end if
                        %>
                    </td>
                </tr>                                   
            </table>
        </div>

        <%
            t.close: set t = nothing
        %>

        <script>
            function filtrar() {
                var sistema = document.getElementById("cboSistema").value;                
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="versiones.asp?s=" + sistema  +"&o=" + ordenamiento;
                window.location.href = vinculo;                      
            }

            function nueva() {
                var vinculo ="crear_version.asp?s=<%= Sistema %>&o=<%= ordenadoPor %>";
                window.location.href = vinculo;
            }    

            function editar(codigo) {
                var vinculo ="editar_version.asp?s=<%= Sistema %>&v=" + codigo + "&o=<%= ordenadoPor %>";
                window.location.href = vinculo;
            }    

            function activar(codigo) {
                var confirmacion = confirm("Desea convertir la versión " + codigo + " en la Versión Activa del Sistema?");                
                var vinculo ="activar_version.asp?s=<%= Sistema %>&v=" + codigo + "&o=<%= ordenadoPor %>";

                if (confirmacion) {     
                    window.location.href = vinculo;
                };                
            }                

            function borrar(codigo) {
                var confirmacion = confirm("Al borrar una version se eliminan los Detalles de la Version \ny el Historial de las actualizaciones de los usuarios\nDesea borrar la version seleccionada?");
                var vinculo ="borrar_version.asp?s=<%= Sistema %>&v=" + codigo + "&o=<%= ordenadoPor %>";                

                if (confirmacion) {     
                    window.location.href = vinculo;
                };
            }         
        </script> 

        <% cc.close: set cc = nothing %>        
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->               
    </body>
</html>