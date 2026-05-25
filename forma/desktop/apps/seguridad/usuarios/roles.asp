<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Asignar Roles a los Usuarios</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->   
        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0025"
            SysLockOut
        %>                    

        <%
            dim cc, t, tt, sqlString, Sistema, ordenadoPor

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")            
        %>    
    </head>

    <body plantilla="lista" reserva="200">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    

        <img src='/imagenes/none.png' style='border: none; width:5px; height: 15px;'>

        <%
            Sistema = request.querystring("s")                    
            ordenadoPor = request.querystring("o")                    

            if Sistema = "" then Sistema = "*"
            if ordenadoPor = "" then ordenadoPor = "0"

            sqlString = "SELECT Codigo, Nombre, Descripcion, TipoRol " & _
                          "FROM (SELECT rolCodigo AS Codigo, rolNombre AS Nombre, rolDescripcion AS Descripcion, " & _
                                      " CASE WHEN TipoRol = 1 THEN 'Rol' ELSE 'Anti-Rol' END AS TipoRol " & _
                                  "FROM seg_Roles "

            if Sistema <> "*" then 
                sqlString = sqlString & "WHERE CodigoSys = '" & Sistema & "'"
            else
                sqlString = sqlString & "WHERE CodigoSys IS NULL"
            end if

            sqlString = sqlString & ") AS r "

            select case ordenadoPor
                case 0: sqlString = sqlString & " ORDER BY Nombre;"
                case 1: sqlString = sqlString & " ORDER BY TipoRol;"
                case 2: sqlString = sqlString & " ORDER BY Nombre DESC;"
                case 3: sqlString = sqlString & " ORDER BY TipoRol DESC;"
            end select

            set t = cc.execute(sqlString)        
        %>

        <div style="width: 98%; margin: auto;">
            <table style="width:100%; margin: auto;">
                <tr style="padding: 10px;">
                    <td colspan ="3" style="text-align:left; width: 20%;">
                        <span style="font-size: 24px">&nbsp;Asignar Roles</span>

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
                        <select name="cboSistema" id="cboSistema" required 
                                class="field"
                                onChange="filtrar()">
                            <option value="*" <% if sistema = "*" then response.write " selected" %>>Multi Sistema</option>    
                            <%
                                set tt = cc.execute("SELECT sysCodigo, sysNombre FROM seg_Sistemas ORDER BY sysNombre;")
                                    if not (tt.bof or tt.eof) then
                                        do
                                            response.write "<option value='" & tt("SysCodigo") & "'"
                                                if tt("sysCodigo") = Sistema then
                                                    response.write " selected"
                                                end if
                                            response.write ">" & tt("sysNombre") & "</option>"

                                            tt.MoveNext
                                        loop until tt.eof
                                    end if
                                tt.close: set tt = nothing
                            %>
                        </select>           

                        &nbsp;             

                        <select name="ordenadoPor" id="ordenadoPor" required 
                                class="field"
                                onChange="filtrar()">
                            <option value="0" <% if ordenadoPor = "0" then response.write " selected" %>>&#9650; Nombre</option>
                            <option value="1" <% if ordenadoPor = "1" then response.write " selected" %>>&#9650; Tipo de Rol</option>
                            <option value="2" <% if ordenadoPor = "2" then response.write " selected" %>>&#9660; Nombre</option>
                            <option value="3" <% if ordenadoPor = "3" then response.write " selected" %>>&#9660; Tipo de Rol</option>
                        </select>                          
                    </td>                    
                </tr>
            </table>

            <table style="width:100%; margin: auto;">
                <tr style="font-size: 14px; background-color: rgb(61, 61, 61); color:rgb(255,255,255);">
                    <td style="padding: 10px; text-align:center; width: 15%;">Tipo</td>
                    <td style="padding: 10px; text-align:center; width: 10%;">Codigo</td>
                    <td style="padding: 10px; text-align:center; width: 30%;">Nombre</td>
                    <td style="padding: 10px; text-align:center; width: 40%;">Descripcion</td>
                    <td style="padding: 10px; text-align:center; width:  5%;">&nbsp;</td>
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
                                            <td style="padding: 15px; text-align:center; width: 15%;" onclick="asignar('<%= t("Codigo") %>')"><%= t("TipoRol") %></td>
                                            <td style="padding: 10px; text-align:  left; width: 10%;" onclick="asignar('<%= t("Codigo") %>')"><%= t("Codigo") %></td>
                                            <td style="padding: 10px; text-align:center; width: 30%;" onclick="asignar('<%= t("Codigo") %>')"><%= t("Nombre") %></td>
                                            <td style="padding: 10px; text-align:center; width: 40%;" onclick="asignar('<%= t("Codigo") %>')"><%= t("Descripcion") %></td>

                                            <td style="padding:  5px; text-align: right; width: 5%;">
                                                <button type="button" class="form-btn azul" onclick="asignar('<%= t("Codigo") %>')">
                                                    <i class=" fa fa-user fa-xl" title="Asignar Rol"></i>
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
                var sistema = document.getElementById("cboSistema").value;
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="roles.asp?s=" + sistema + "&o=" + ordenamiento;
                window.location.href = vinculo;                      
            }

            function asignar(codigo) {
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var vinculo ="roles_asignar.asp?r=" + codigo + "&o=" + ordenamiento;
                window.location.href = vinculo;
            }          
        </script> 

        <% cc.close: set cc = nothing %>     
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>