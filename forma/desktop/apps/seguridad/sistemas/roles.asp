<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Roles de los Sistemas</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->            
        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0080"
            SysLockOut

            dim cc, t, tt, sqlString, ordenadoPor
            dim Sistema

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")            
        %>    
    </head>

    <body plantilla="lista" reserva="200">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    

        <img src='imagenes/none.png' style='border: none; width:5px; height: 15px;'>

        <%
            ordenadoPor = request.querystring("o")                    
            if ordenadoPor = "" then ordenadoPor = "0"

            sqlString = "SELECT Codigo, Nombre, Descripcion, TipoRol " & _
                          "FROM (SELECT rolCodigo AS Codigo, rolNombre AS Nombre, rolDescripcion AS Descripcion, " & _
                                      " CASE WHEN TipoRol = 1 THEN 'Rol' ELSE 'Anti-Rol' END AS TipoRol " & _
                                  "FROM seg_Roles " & _
                                 "WHERE (CodigoSys IS NULL) " & _
                               ") AS r "

            select case ordenadoPor
                case 0: sqlString = sqlString & " ORDER BY Nombre;"
                case 1: sqlString = sqlString & " ORDER BY TipoRol;"
                case 2: sqlString = sqlString & " ORDER BY Nombre DESC;"
                case 3: sqlString = sqlString & " ORDER BY TipoRol DESC;"
            end select

            set t = cc.execute(sqlString)        
        %>

        <div style="width:98%; margin: auto;">
            <table style="width:100%; margin: auto;">
                <tr class="noborder" style="padding: 10px;">
                    <td colspan ="3" style="text-align:left; width: 20%;">
                        <span style="font-size: 24px">&nbsp;Roles Múltiples</span>

                        <br />

                        <span style="font-size: 20px">
                            &nbsp;Ordenados por&nbsp;
                            <%
                                select case ordenadoPor
                                    case 0: response.write "Nombre"
                                    case 1: response.write "Tipo de Rol"
                                    case 2: response.write "Nombre (descendentemente)"
                                    case 3: response.write "Tipo de Rol (descendentemente)"
                                end select                            
                            %>
                        </span>
                    </td>

                    <td colspan ="2" style="text-align:right; width: 20%;">
                        <select name="ordenadoPor" id="ordenadoPor" required 
                                class="field"
                                onChange="filtrar()">
                            <option value="0" <% if ordenadoPor = "0" then response.write " selected" %>>&#9650; Nombre</option>
                            <option value="1" <% if ordenadoPor = "1" then response.write " selected" %>>&#9650; Tipo de Rol</option>
                            <option value="2" <% if ordenadoPor = "2" then response.write " selected" %>>&#9660; Nombre</option>
                            <option value="3" <% if ordenadoPor = "3" then response.write " selected" %>>&#9660; Tipo de Rol</option>
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
                    <td style="padding: 10px; text-align:center; width: 15%;">Tipo</td>
                    <td style="padding: 10px; text-align:center; width: 10%;">Codigo</td>
                    <td style="padding: 10px; text-align:center; width: 30%;">Nombre</td>
                    <td style="padding: 10px; text-align:center; width: 35%;">Descripcion</td>
                    <td style="padding: 10px; text-align:center; width: 10%;">&nbsp;</td>
                </tr>

                <tr>
                    <td colspan="5">
                        <div id="overFlow" style="width:100%; height: 625px; overflow: auto; background-color: rgb(207, 207, 207);">                        
                            <table style="width: 100%;">
                                <%  
                                    if not (t.bof or t.eof) then  
                                        conta = 0

                                        Do     
                                            conta = conta + 1
                                %>
                                        <tr style="font-size: 14px; <%
                                                if t("TipoRol") = "Rol" then
                                                    response.write "background-color: rgb(255,255,255);"
                                                else
                                                    response.write "background-color: rgb(255, 222, 228);"
                                                end if
                                            %> color:rgb(0,0,0); border-bottom: 1px solid rgb(194, 194, 194);">
                                            <td style="padding: 15px; text-align:center; width: 15%;" onclick="editar('<%= t("Codigo") %>')"><%= t("TipoRol") %></td>
                                            <td style="padding: 10px; text-align:  left; width: 10%;" onclick="editar('<%= t("Codigo") %>')"><%= t("Codigo") %></td>
                                            <td style="padding: 10px; text-align:center; width: 30%;" onclick="editar('<%= t("Codigo") %>')"><%= t("Nombre") %></td>
                                            <td style="padding: 10px; text-align:center; width: 35%;" onclick="editar('<%= t("Codigo") %>')"><%= t("Descripcion") %></td>

                                            <td style="padding:  5px; text-align: right; width: 10%;">
                                                <button type="button" class="form-btn azul" onclick="sUsuarios('<%= t("Codigo") %>')">
                                                    <i class=" fa fa-user fa-xl" title="Asignar Roles"></i>
                                                </button>        

                                                <button type="button" class="form-btn rojo" onclick="borrar('<%= t("Codigo") %>')">
                                                    <i class=" fa fa-trash fa-xl" title="Borrar Rol"></i>
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
                                response.write "No se encontró ning{un rol"
                            else
                                if conta = 1 then
                                    response.write "Se encontró un rol"
                                else
                                    response.write "Se encontraron " & conta & " roles"
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

                var vinculo ="roles.asp?o=" + ordenamiento;
                window.location.href = vinculo;                      
            }
                  
            function editar(codigo) {
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var vinculo ="editar_rol.asp?r=" + codigo + "&o=" + ordenamiento;
                window.location.href = vinculo;
            }    

            function borrar(Rol) {
                var confirmacion = confirm("Desea borrar el Rol seleccionada?");
                var vinculo ="borrar_rol.asp?r=" + Rol + "&o=<%= ordenadoPor %>";                

                if (confirmacion) {     
                    window.location.href = vinculo;
                };
            }    

            function sUsuarios(codigo) {
                var vinculo ="/forma/desktop/apps/seguridad/usuarios/roles_asignar.asp?r=" + codigo + "&o=1";               
                window.location.href = vinculo;                
            }   
        </script> 

        <% cc.close: set cc = nothing %>        
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->               
    </body>
</html>