<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Procesos de los Sistemas</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->    
        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0070"
            SysLockOut
        %>

        <%
            dim cc, t, tt, sqlString, ordenadoPor, conta
            dim Sistema

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
        %>    
    </head>

    <body plantilla="Lista">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    

        <img src='imagenes/none.png' style='border: none; width:5px; height: 15px;'>

        <%
            Sistema = Request.querystring("s")
            ordenadoPor = request.querystring("o")           
         
            if Sistema = "" then Sistema = SistemaDefault()
            if ordenadoPor = "" then ordenadoPor = "0"

            sqlString = "SELECT Sistema, Proceso, Estatus, Tipo, Home, Snip, Shortcut, Movil, Nombre, Indice, Icon, PerteneceA " & _
                          "FROM (SELECT p.proSistema AS Sistema, p.proCodigo AS Proceso, CASE WHEN p.proActivo = 1 THEN 'Activo' ELSE 'Inactivo' END AS Estatus, " & _
                                      " p.proMenuItem AS Tipo, p.proHomePage AS Home, CASE WHEN p.Snippet IS NULL THEN 0 ELSE 1 END AS Snip, p.Shortcut, " & _
                                      " p.proNombre AS Nombre, p.proMenuIndice AS Indice, p.proIcon AS Icon, Pa.proNombre AS PerteneceA, p.Movil " & _
                                  "FROM dbo.seg_Procesos AS p " & _
                       "LEFT OUTER JOIN dbo.seg_Procesos AS Pa " & _
                                    "ON p.proSistema = Pa.proSistema " & _
                                   "AND p.proRoot = Pa.proCodigo " & _
                                 "WHERE (p.proSistema = '" & Sistema & "') " & _
                              " ) AS Procs "

            select case ordenadoPor
                case 0: sqlString = sqlString & " ORDER BY Proceso;"
                case 1: sqlString = sqlString & " ORDER BY Nombre;"
                case 2: sqlString = sqlString & " ORDER BY Estatus;"
                case 3: sqlString = sqlString & " ORDER BY PerteneceA, Indice ASC;"
                case 4: sqlString = sqlString & " ORDER BY Proceso DESC;"
                case 5: sqlString = sqlString & " ORDER BY Nombre DESC;"
                case 6: sqlString = sqlString & " ORDER BY Estatus DESC;"
                case 7: sqlString = sqlString & " ORDER BY PerteneceA DESC, Indice ASC;"
            end select

            set t = cc.execute(sqlString)        
        %>

        <div style="width:98%; margin: auto;">
            <table style="width:100%; margin: auto;">
                <tr class="noborder" style="padding: 10px;">
                    <td colspan ="3" style="text-align:left; width: 20%;">
                        <span style="font-size: 24px">
                            &nbsp;Procesos - <%= NombreSistema(Sistema) %>
                        </span>

                        <br />

                        <span style="font-size: 20px">
                            <%
                                select case ordenadoPor
                                    case 0: response.write "&nbsp;Ordenada por Proceso"
                                    case 1: response.write "&nbsp;Ordenada por Nombre"
                                    case 2: response.write "&nbsp;Ordenada por Estado"
                                    case 3: response.write "&nbsp;Ordenada por Pertenece A"
                                    case 4: response.write "&nbsp;Ordenada por Proceso (descendentemente)"
                                    case 5: response.write "&nbsp;Ordenada por Nombre (descendentemente)"
                                    case 6: response.write "&nbsp;Ordenada por Estatus (descendentemente)"
                                    case 7: response.write "&nbsp;Ordenada por Pertenece A (descendentemente)"
                                end select                            
                            %>
                        </span>
                    </td>

                    <td colspan ="4" style="text-align:right; width: 20%;">
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
                            <option value="0" <% if ordenadoPor = "0" then response.write " selected" %>>&#9650; Proceso</option>
                            <option value="1" <% if ordenadoPor = "1" then response.write " selected" %>>&#9650; Nombre</option>
                            <option value="2" <% if ordenadoPor = "2" then response.write " selected" %>>&#9650; Estado</option>
                            <option value="3" <% if ordenadoPor = "3" then response.write " selected" %>>&#9650; Pertenece A</option>
                            <option value="4" <% if ordenadoPor = "4" then response.write " selected" %>>&#9660; Proceso</option>
                            <option value="5" <% if ordenadoPor = "5" then response.write " selected" %>>&#9660; Nombre</option>
                            <option value="6" <% if ordenadoPor = "6" then response.write " selected" %>>&#9660; Estado</option>
                            <option value="7" <% if ordenadoPor = "7" then response.write " selected" %>>&#9660; Pertenece A</option>
                        </select>         

                        &nbsp;&nbsp;

                        <button type="button" class="form-btn verde" onclick="editar('*')">
                            <i class=" fa fa-edit fa-xl" title="Nuevo"></i>
                        </button>                        
                    </td>                    
                </tr>
            </table>            

            <table style="width:100%; margin: auto;">
                <tr style="font-size: 14px; background-color: rgb(61, 61, 61); color:rgb(255,255,255);">
                    <td style="padding: 10px; text-align:center; width: 15%;">Proceso</td>
                    <td style="padding: 10px; text-align:center; width: 25%;">Tipo</td>
                    <td style="padding: 10px; text-align:left; width: 20%;">Nombre</td>
                    <td style="padding: 10px; text-align:center; width:  5%;">Indice</td>
                    <td style="padding: 10px; text-align:center; width: 20%;">Pertenece A</td>
                    <td style="padding: 10px; text-align:center; width: 15%;">&nbsp;</td>
                </tr>

                <tr>
                    <td colspan="6">
                        <div id="overFlow" style="width:100%; height: 625px; overflow: auto; background-color: rgb(207, 207, 207);">                        
                            <table style="width: 100%;">
                                <%  
                                    if not (t.bof or t.eof) then  
                                        conta = 0
                                        Do     
                                            conta = conta + 1 
                                %>
                                        <tr style="font-size: 14px; <%
                                                if t("Estatus") = "Activo" then
                                                    response.write "background-color: rgb(255,255,255);"
                                                else
                                                    response.write "background-color: rgb(255, 222, 228);"
                                                end if
                                            %> color:rgb(0,0,0); border-bottom: 1px solid rgb(194, 194, 194);">
                                            <td style="padding: 10px; text-align: center; width: 15%;" onclick="editar('<%= t("Proceso") %>')"><%= t("Proceso") %></td>                                            
                                            <td style="padding:  5px; text-align: center; width: 25%;" onclick="editar('<%= t("Proceso") %>')">
                                                <%
                                                    ContarEstatus = 1

                                                    if t("Tipo") = 0 then 
                                                        response.write "<img src='imagenes/menu.png' style='border: none; width: 25px; height: 25px;'>&nbsp;&nbsp;"
                                                    else
                                                        response.write "<img src='imagenes/menu_item.png' style='border: none; width: 25px; height: 25px;'>&nbsp;&nbsp;"
                                                    end if

                                                    if t("Home") = 1 then 
                                                        ContarEstatus = ContarEstatus + 1
                                                        response.write "<img src='imagenes/home_page.png' style='border: none; width: 25px; height: 25px;'>&nbsp;&nbsp;" 
                                                    end if

                                                    if t("Snip") = 1  then
                                                        ContarEstatus = ContarEstatus + 1
                                                        response.write "<img src='imagenes/snippet.png' style='border: none; width: 25px; height: 25px;'>&nbsp;&nbsp;" 
                                                    end if   

                                                    if t("Shortcut") = 1 then 
                                                        ContarEstatus = ContarEstatus + 1
                                                        response.write "<img src='imagenes/vinculo.png' style='border: none; width: 25px; height: 25px;'>&nbsp;&nbsp;" 
                                                    end if     
                                                    
                                                    if t("Movil") = 1 then 
                                                        ContarEstatus = ContarEstatus + 1
                                                        response.write "<img src='imagenes/movil.png' style='border: none; width: 25px; height: 25px;'>&nbsp;&nbsp;" 
                                                    end if  

                                                    for contK = (ContarEstatus + 1) to 5
                                                        response.write "<img src='imagenes/empty.png' style='border: none; width: 25px; height: 25px;'>&nbsp;&nbsp;"
                                                    next
                                                %>
                                            </td>
                                            <td style="padding: 10px; text-align:left; width: 20%;" onclick="editar('<%= t("Proceso") %>')"><%= t("Nombre") %></td>
                                            <td style="padding: 10px; text-align:center; width:  5%;" onclick="editar('<%= t("Proceso") %>')"><%= t("Indice") %></td>
                                            <td style="padding: 10px; text-align:center; width: 20%;" onclick="editar('<%= t("Proceso") %>')"><%= t("PerteneceA") %></td>

                                            <td style="padding:  5px; text-align: right; width: 15%;">
                                                <button type="button" class="form-btn azul"  onclick="asignar('<%= Sistema %>', '<%= t("Proceso") %>')">
                                                    <i class=" fa fa-user fa-xl" title="Asignar Roles"></i>
                                                </button>

                                                <button type="button" class="form-btn rojo"  onclick="borrar('<%= Sistema %>','<%= t("Proceso") %>')">
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
                                response.write "No se encontró ningún proceso"
                            else
                                if conta = 1 then
                                    response.write "Se encontró un proceso"
                                else
                                    response.write "Se encontraron " & conta & " procesos"
                                end if
                            end if
                        %>
                    </td>
                </tr>                               
            </table>
        </div>

        <script>
            function filtrar() {
                var sistema = document.getElementById("cboSistema").value;
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="procesos.asp?s=" + sistema + "&o=" + ordenamiento;
                window.location.href = vinculo;                      
            }
                  
            function editar(codigo) {
                var sistema = document.getElementById("cboSistema").value;
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="editar_proceso.asp?s=" + sistema + "&p=" + codigo + "&op=" + ordenamiento;
                window.location.href = vinculo;
            }    

            function borrar(Sistema, Proceso) {
                var confirmacion = confirm("Desea borrar el proceso seleccionado?");
                var vinculo ="borrar_proceso.asp?s=" + Sistema + "&p=" + Proceso + "&o=<%= ordenadoPor %>";

                if (confirmacion) {     
                    window.location.href = vinculo;
                };
            }    

            function asignar(sistema, proceso) {
                var vinculo ="/forma/desktop/apps/seguridad/usuarios/permisos_asignar.asp?s=" + sistema + "&p=" + proceso + "&o=1";
                window.location.href = vinculo;
            }      
        </script> 

        <% 
            t.close: set t = nothing 
            cc.close: set cc = nothing 
        %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->    	                   
    </body>
</html>