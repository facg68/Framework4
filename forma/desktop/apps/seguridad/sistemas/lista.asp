<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Sistemas</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->            
        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0060"
            SysLockOut
        %>     

        <%
            dim cc, t, tt, sqlString, data, labels, conta
            dim cActivas, cInactivas, estatusAnuncio, ordenadoPor

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn") 
        %>    
    </head>

    <body plantilla="lista" reserva="200">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    

        <br />

        <%
            ordenadoPor = request.querystring("op")            
            if ordenadoPor = "" then ordenadoPor = "0"

            sqlString = "SELECT Codigo, Nombre, Descripcion, VersionActual, Cambios, CambiosTotales, Roles " & _
                        "FROM seg_SistemasVersiones " 

            select case ordenadoPor
                case 0: sqlString = sqlString & " ORDER BY Nombre;"
                case 1: sqlString = sqlString & " ORDER BY Descripcion;"
                case 2: sqlString = sqlString & " ORDER BY VersionActual;"
                case 3: sqlString = sqlString & " ORDER BY Cambios"
                case 4: sqlString = sqlString & " ORDER BY CambiosTotales"
                case 5: sqlString = sqlString & " ORDER BY Nombre DESC;"
                case 6: sqlString = sqlString & " ORDER BY Descripcion DESC;"
                case 7: sqlString = sqlString & " ORDER BY VersionActual DESC;"
                case 8: sqlString = sqlString & " ORDER BY sysMenuIndice DESC;"
                case 9: sqlString = sqlString & " ORDER BY CambiosTotales DESC;"
            end select

            set t = cc.execute(sqlString)        
        %>

        <div style="width: 98%; margin: auto;">
            <table style="width:100%;">
                <tr class="noborder" style="padding: 10px;">
                    <td style="text-align:left; width: 70%;">
                        <span style="font-size: 24px">
                            &nbsp;Lista de Sistemas
                        </span>

                        <br />

                        <span style="font-size: 20px">
                            <%
                                select case ordenadoPor
                                    case 0: response.write "&nbsp;Ordenada por Nombre"
                                    case 1: response.write "&nbsp;Ordenada por Descripcion"
                                    case 2: response.write "&nbsp;Ordenada por VersionActual"
                                    case 3: response.write "&nbsp;Ordenada por Cambios"
                                    case 4: response.write "&nbsp;Ordenada por CambiosTotales"
                                    case 5: response.write "&nbsp;Ordenada por Nombre (descendentemente)"
                                    case 6: response.write "&nbsp;Ordenada por Descripcion (descendentemente)"
                                    case 7: response.write "&nbsp;Ordenada por VersionActual (descendentemente)"
                                    case 8: response.write "&nbsp;Ordenada por Cambios (descendentemente)"
                                    case 9: response.write "&nbsp;Ordenada por CambiosTotales (descendentemente)"
                                end select                            
                            %>
                        </span>
                    </td>

                    <td style="text-align:right; width: 30%;">
                        <select name="ordenadoPor" id="ordenadoPor" required 
                                class="field"
                                onChange="filtrar()">
                            <option value="0" <% if ordenadoPor = "0" then response.write " selected" %>>&#9650; Nombre</option>
                            <option value="1" <% if ordenadoPor = "1" then response.write " selected" %>>&#9650; Descripcion</option>
                            <option value="2" <% if ordenadoPor = "2" then response.write " selected" %>>&#9650; Version Actual</option>
                            <option value="3" <% if ordenadoPor = "3" then response.write " selected" %>>&#9650; Cambios</option>
                            <option value="4" <% if ordenadoPor = "4" then response.write " selected" %>>&#9650; Cambios Totales</option>
                            <option value="5" <% if ordenadoPor = "5" then response.write " selected" %>>&#9660; Nombre</option>
                            <option value="6" <% if ordenadoPor = "6" then response.write " selected" %>>&#9660; Descripcion</option>
                            <option value="7" <% if ordenadoPor = "7" then response.write " selected" %>>&#9660; Version Actual</option>
                            <option value="8" <% if ordenadoPor = "8" then response.write " selected" %>>&#9660; Cambios</option>
                            <option value="9" <% if ordenadoPor = "9" then response.write " selected" %>>&#9660; Cambios Totales</option>
                        </select>         

                        &nbsp;

                        <button type="button" class="form-btn verde" onclick="editar('*')">
                            <i class=" fa fa-edit fa-xl" title="Nuevo"></i>
                        </button>                        
                    </td>                    
                </tr>
            </table>

            <table style="width:98%;">
                <tr style="font-size: 14px; background-color: rgb(61, 61, 61); color:rgb(255,255,255);">
                    <td style="padding: 10px; text-align:center; width: 10%;">Codigo</td>
                    <td style="padding: 10px; text-align:center; width: 15%;">Nombre</td>
                    <td style="padding: 10px; text-align:center; width: 15%;">Descripcion</td>
                    <td style="padding: 10px; text-align:center; width: 10%;">Roles</td>
                    <td style="padding: 10px; text-align:center; width: 10%;">Cambios</td>
                    <td style="padding: 10px; text-align:center; width: 10%;">C.Totales</td>
                    <td style="padding: 10px; text-align:center; width: 30%;">&nbsp;</td>
                </tr>

                <tr>
                    <td colspan="12">
                        <div id="overFlow" style="width:100%; height: 650px; overflow: auto; background-color: rgb(207, 207, 207);">                        
                            <table style="width: 100%;">
                                <%  
                                    if not (t.bof or t.eof) then  
                                        conta = 0
                                        Do     
                                            conta = conta + 1
                                                                            %>
                                        <tr style="font-size: 14px; background-color: rgb(255,255,255); color:rgb(0,0,0); border-bottom: 1px solid rgb(194, 194, 194);">
                                            <td style="padding: 10px; text-align:center; width: 10%;" onclick="editar('<%= t("Codigo") %>')"><%= t("Codigo") %></td>
                                            <td style="padding: 10px; text-align:center; width: 15%;" onclick="editar('<%= t("Codigo") %>')"><%= t("Nombre") %></td>
                                            <td style="padding: 10px; text-align:center; width: 15%;" onclick="editar('<%= t("Codigo") %>')"><%= t("Descripcion") %></td>
                                            <td style="padding: 10px; text-align:center; width: 10%;" onclick="editar('<%= t("Codigo") %>')"><%= t("Roles") %></td>
                                            <td style="padding: 10px; text-align:center; width: 10%;" onclick="editar('<%= t("Codigo") %>')"><%= t("Cambios") %></td>
                                            <td style="padding: 10px; text-align:center; width: 10%;" onclick="editar('<%= t("Codigo") %>')"><%= t("CambiosTotales") %></td>

                                            <td style="padding: 10px; text-align:right; width: 30%;">
                                                <button type="button" class="form-btn azul" style="background-color: rgb(102, 153, 255); color: white;" onclick="sProcesos('<%= t("Codigo") %>')">
                                                    <i class=" fa fa-gears fa-xl" title="Procesos"></i>
                                                </button>

                                                <button type="button" class="form-btn " style="background-color: rgb(179, 128, 89); color: white;" onclick="sVariables('<%= t("Codigo") %>')">
                                                    <i class=" fa fa-font fa-xl" title="Variables"></i>
                                                </button>

                                                <button type="button" class="form-btn " style="background-color: rgb(144, 111, 197); color: white;" onclick="sVersiones('<%= t("Codigo") %>')">
                                                    <i class=" fa fa-code-fork fa-xl" title="Versiones"></i>
                                                </button>

                                                <% if t("Roles") = 0 then %>
                                                    <button type="button" class="form-btn "style="background-color: rgb(30, 96, 106); color: white;" disabled>
                                                        <i class=" fa fa-sliders fa-xl" title="Roles"></i>
                                                    </button>
                                                <% else %>
                                                    <button type="button" class="form-btn "style="background-color: rgb(30, 96, 106); color: white;" onclick="sRoles('<%= t("Codigo") %>')">
                                                        <i class=" fa fa-sliders fa-xl" title="Roles"></i>
                                                    </button>
                                                <% end if %>                                                

                                                <button type="button" class="form-btn " style="background-color: rgb(168, 60, 60); color: white;" onclick="borrar('<%= t("Codigo") %>')">
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
                    <td colspan="7" style="padding: 10px; text-align:center; width: 100%;">
                        <% 
                            if conta = 0 then   
                                response.write "No se encontró ningún sistema"
                            else
                                if conta = 1 then
                                    response.write "Se encontró un sistema"
                                else
                                    response.write "Se encontraron " & conta & " sistemas"
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
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="lista.asp?op=" + ordenamiento;
                window.location.href = vinculo;                      
            }
                  
            function editar(codigo) {
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var vinculo ="editar_sistema.asp?s=" + codigo + "&op=" + ordenamiento;
                window.location.href = vinculo;
            }    

            function sProcesos(codigo) {
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var vinculo ="procesos.asp?s=" + codigo + "&op=" + ordenamiento;
                window.location.href = vinculo;
            }      

            function sVariables(codigo) {
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var vinculo ="variables.asp?s=" + codigo + "&op=" + ordenamiento;
                window.location.href = vinculo;
            }                 
            
            function sVersiones(codigo) {
                var vinculo ="versiones.asp?s=" + codigo;
                window.location.href = vinculo;
            }      
            
            function sRoles(codigo) {
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var vinculo ="roles_sys.asp?s=" + codigo + "&o=" + ordenamiento ;
                window.location.href = vinculo;                
            }

            function borrar(codigo) {
                var confirmacion = confirm("Desea borrar el Sistema seleccionado?");
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var vinculo ="borrar_sistema.asp?s=" + codigo + "&op=" + ordenamiento;

                if (confirmacion) {     
                    window.location.href = vinculo;
                };
            }         
        </script> 

        <% cc.close: set cc = nothing %>        
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->    	               
    </body>
</html>