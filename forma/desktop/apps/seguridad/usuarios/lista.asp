<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Usuarios de los Sistemas</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  

        <style>
            a.letra, a.letra:hover, a.letra:focus {
                color: #000;
                text-decoration: none;
                outline: 0;
            }            
        </style>

        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0020"
            SysLockOut

            dim cc, t, tt, sqlString, ordenadoPor, letra
            dim Sistema

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")    

            Sub PresentarIniciales    
                dim ini, comando, vinculo

                comando = "SELECT Inicial FROM (SELECT LEFT(usuNombre, 1) AS Inicial, COUNT(*) AS Cuantos " & _
                                                 "FROM dbo.seg_Usuarios WHERE (usuNombre IS NOT NULL) AND (usuCodigo <> 'defaults') " & _
                                             "GROUP BY LEFT(usuNombre, 1)) AS t;"

                set ini = cc.execute(comando)
                    if not (ini.bof or ini.eof) then
                        Do
                            vinculo ="#" & uCase(ini("Inicial"))
                            response.write "<a class='letra' href='" & vinculo & "'>" & uCase(ini("Inicial")) & "</a>&nbsp;&nbsp;"
                            ini.MoveNext
                        Loop Until (ini.eof)
                    end if
                ini.close: set ini = nothing
            End Sub
        %>    
    </head>

    <body plantilla="lista" reserva="200">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    

        <br />

        <%
            ordenadoPor = request.querystring("o")                    
            if ordenadoPor = "" then ordenadoPor = "0"

            sqlString = "SELECT usuCodigo, usuNombre, usuCargo, usuCorreo, usuTelefono, " & _
                          "CASE WHEN usuEstado = 1 THEN 'Activo' ELSE 'Desactivado' END AS Estatus " & _
                          "FROM seg_Usuarios " & _
                         "WHERE (usuCodigo <> 'DEFAULTS') "

            select case ordenadoPor
                case 0: sqlString = sqlString & " ORDER BY usuNombre;"
                case 1: sqlString = sqlString & " ORDER BY usuCodigo;"
                case 2: sqlString = sqlString & " ORDER BY usuNombre DESC;"
                case 3: sqlString = sqlString & " ORDER BY usuCodigo DESC;"
            end select

            set t = cc.execute(sqlString)        
        %>

        <div style="width: 98%; margin: auto;">
            <table style="width:100%; margin: auto;">
                <tr style="padding: 10px;">
                    <td colspan ="3" style="text-align:left; width: 20%;">
                        <span style="font-size: 24px">&nbsp;Usuarios</span>

                        <br />

                        <span style="font-size: 20px">
                            &nbsp;Ordenados por&nbsp;
                            <%
                                select case ordenadoPor
                                    case 0: response.write "Nombre"
                                    case 1: response.write "Codigo"
                                    case 2: response.write "Nombre (descendentemente)"
                                    case 3: response.write "Codigo (descendentemente)"
                                end select                            
                            %>
                        </span>
                    </td>

                    <td colspan ="4" style="text-align:right; width: 20%;">
                        <% PresentarIniciales %>
                        <select name="ordenadoPor" id="ordenadoPor" required 
                                class="field"
                                onChange="filtrar()">
                            <option value="0" <% if ordenadoPor = "0" then response.write " selected" %>>&#9650; Nombre</option>
                            <option value="1" <% if ordenadoPor = "1" then response.write " selected" %>>&#9650; Codigo</option>
                            <option value="2" <% if ordenadoPor = "2" then response.write " selected" %>>&#9660; Nombre</option>
                            <option value="3" <% if ordenadoPor = "3" then response.write " selected" %>>&#9660; Codigo</option>
                        </select>         

                        &nbsp;&nbsp;

                        <button type="button" class="form-btn verde"  onclick="crear_usuario()">
                            <i class=" fa fa-edit fa-xl" title="Nuevo"></i>
                        </button>                        
                    </td>                    
                </tr>
            </table>

            <table style="width:100%; margin: auto;">            
                <tr style="font-size: 14px; background-color: rgb(61, 61, 61); color:rgb(255,255,255);">
                    <td style="padding: 10px; text-align:center; width: 10%;">Codigo</td>
                    <td style="padding: 10px; text-align:center; width: 20%;">Nombre</td>
                    <td style="padding: 10px; text-align:center; width: 20%;">Cargo</td>
                    <td style="padding: 10px; text-align:center; width: 20%;">Correo</td>
                    <td style="padding: 10px; text-align:center; width: 30%;">&nbsp;</td>
                </tr>

                <tr>
                    <td colspan="5">
                        <div id="overFlow" style="width:100%; height: 625px; overflow: auto; background-color: rgb(207, 207, 207);">                        
                            <table style="width: 100%;">
                                <%  
                                    if not (t.bof or t.eof) then  
                                        conta = 0
                                        letra = "*/"

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

                                            <td style="padding: 15px; text-align:center; width: 10%;" onclick="editar('<%= t("usuCodigo") %>')">
                                            <%
                                                if left(t("usuNombre"), 1) <> letra then
                                                    letra = left(t("usuNombre"), 1)
                                                    response.write "<a class='letra' name='" & left(t("usuNombre"), 1) & "' >" & t("usuCodigo") & "</a>"
                                                else
                                                    response.write t("usuCodigo")
                                                end if                                            
                                            %>                                            
                                            </td>
                                            <td style="padding: 10px; text-align:  left; width: 20%;" onclick="editar('<%= t("usuCodigo") %>')"><%= t("usuNombre") %></td>
                                            <td style="padding: 10px; text-align:  left; width: 20%;" onclick="editar('<%= t("usuCodigo") %>')"><%= t("usuCargo") %></td>
                                            <td style="padding: 10px; text-align:  left; width: 20%;" onclick="editar('<%= t("usuCodigo") %>')"><%= t("usuCorreo") %></td>

                                            <td style="padding:  5px; text-align: right; width: 30%;">

                                                <button type="button" class="form-btn naranja" 
                                                        onclick="usu_variables('<%= t("usuCodigo") %>', '<%= t("usuNombre") %>')"
                                                            <%
                                                                if PuedeAccesar("seguridad", "seg.0040") = 0 then
                                                                    response.write " disabled"
                                                                end if
                                                            %>
                                                        >
                                                    <i class=" fa fa-font fa-xl" title="Asignar Variables"></i>
                                                </button>

                                                <button type="button" class="form-btn violeta" 
                                                        onclick="usu_roles('<%= t("usuCodigo") %>', '<%= t("usuNombre") %>')"
                                                            <%
                                                                if PuedeAccesar("seguridad", "seg.0025") = 0 then
                                                                    response.write " disabled"
                                                                end if
                                                            %>                                                        
                                                        >
                                                    <i class=" fa fa-gear fa-xl" title="Asignar Roles"></i>
                                                </button>                                                

                                                <button type="button" class="form-btn azul" 
                                                        onclick="verDatos('<%= t("usuCodigo") %>', '<%= t("usuNombre") %>')">
                                                    <i class=" fa fa-list fa-xl" title="Datos del Usuario"></i>
                                                </button>

                                                <button type="button" class="form-btn" 
                                                        style="background-color: rgb(225, 161, 49);" 
                                                        onclick="usuMenu('<%= t("usuCodigo") %>', '<%= t("usuNombre") %>')">
                                                    <i class=" fa fa-exclamation fa-xl" title="Re-crear menu del Usuario"></i>
                                                </button>                                                

                                                <button type="button" class="form-btn rojo"
                                                        onclick="borrar('<%= t("usuCodigo") %>', '<%= t("usuNombre") %>')">
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
                    <td colspan="5" style="padding: 10px; text-align:center; width: 100%;">
                        <% 
                            if conta = 0 then   
                                response.write "No se encontró ningun usuario"
                            else
                                if conta = 1 then
                                    response.write "Se encontró un usuario"
                                else
                                    response.write "Se encontraron " & conta & " usuarios"
                                end if
                            end if
                        %>
                    </td>
                </tr>                               
            </table>
        </div>

        <% t.close: set t = nothing %>

        <script>
            function filtrar() {
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="lista.asp?o=" + ordenamiento;
                window.location.href = vinculo;                      
            }
                  
            function editar(codigo) {
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var vinculo ="editar_usuario.asp?u=" + codigo + "&o=" + ordenamiento;
                window.location.href = vinculo;
            }    

            function crear_usuario() {
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var vinculo ="crear_usuario.asp?o=" + ordenamiento;
                window.location.href = vinculo;                
            }

            function borrar(usuario, nombre) {
                var confirmacion = confirm("Desea borrar el usuario " + nombre + "?");
                var vinculo ="borrar_usuario.asp?u=" + usuario + "&o=<%= ordenadoPor %>";                

                if (confirmacion) {     
                    window.location.href = vinculo;
                };
            }    

            function verDatos(usuario) {
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var vinculo ="datos_usuario.asp?u=" + usuario + "&o=" + ordenamiento;
                window.location.href = vinculo;                
            }

            function usu_roles(usuario, nombre) {
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var vinculo ="editar_usuario_roles.asp?u=" + usuario + "&o=" + ordenamiento;
                window.location.href = vinculo;                
            }

            function usu_variables(usuario, nombre) {
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var vinculo ="editar_usuario_variables.asp?u=" + usuario + "&o=" + ordenamiento;
                window.location.href = vinculo;                
            }            

            function usuMenu(usuario, nombre) {
                var confirmacion = confirm("Desea Re-Crear el Menu del Usuario " + nombre + "?");
                var vinculo ="menu_usuario.asp?u=" + usuario + "&o=<%= ordenadoPor %>";                

                if (confirmacion) {     
                    window.location.href = vinculo;
                };                            
            }       
        </script> 

        <% cc.close: set cc = nothing %>
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>