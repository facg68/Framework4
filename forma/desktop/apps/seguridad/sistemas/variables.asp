<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Variables Globales</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->            
        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0085"
            SysLockOut

            dim cc, t, tt, sqlString, ordenadoPor

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

    <body plantilla="lista" reserva="225">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    

        <br />

        <%
            Sistema = Request.querystring("s")
            ordenadoPor = request.querystring("o")           
         
            if Sistema = "" then Sistema = SistemaDefault()
            if ordenadoPor = "" then ordenadoPor = "2"

            sqlString = "SELECT Parametro, Descripcion, Exponer FROM seg_Parametros WHERE (Sistema = '" & sistema & "') "

            select case ordenadoPor
                case 1: sqlString = sqlString & " ORDER BY Parametro;"
                case 2: sqlString = sqlString & " ORDER BY Descripcion;"
                case 3: sqlString = sqlString & " ORDER BY Parametro DESC;"                
                case 4: sqlString = sqlString & " ORDER BY Descripcion DESC;"                
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
                                    <span style="font-size: 24px">&nbsp;Variables  - <%= NombreSistema(Sistema) %></span>

                                    <br />

                                    <span style="font-size: 20px">
                                        &nbsp;Ordenados por&nbsp;
                                        <%
                                            select case ordenadoPor
                                                case 1: response.write "Nombre"
                                                case 2: response.write "Descripción"
                                                case 3: response.write "Nombre (descendentemente)"
                                                case 4: response.write "Descripción (descendentemente)"
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

                                    &nbsp;

                                    <select name="ordenadoPor" id="ordenadoPor" required 
                                            class="field"
                                            onChange="filtrar()">
                                        <option value="1" <% if ordenadoPor = "1" then response.write " selected" %>>&#9650; Nombre</option>
                                        <option value="2" <% if ordenadoPor = "2" then response.write " selected" %>>&#9650; Descripcion</option>
                                        <option value="3" <% if ordenadoPor = "3" then response.write " selected" %>>&#9660; Nombre</option>
                                        <option value="4" <% if ordenadoPor = "4" then response.write " selected" %>>&#9660; Descripcion</option>
                                    </select>         

                                    &nbsp;&nbsp;

                                    <button type="button" class="form-btn verde" onclick="editar('<%= Sistema %>', '*')">
                                        <i class=" fa fa-edit fa-xl" title="Nueva"></i>
                                    </button>                                 
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>

            <table>
                <tr style="font-size: 14px; background-color: rgb(61, 61, 61); color:rgb(255,255,255);">
                    <td>
                        <table style="width: 100%;">                
                            <tr style="font-size: 14px; background-color: rgb(61, 61, 61); color:rgb(255,255,255);">
                                <td style="padding: 10px; text-align: center; width: 15%">Nombre</td>
                                <td style="padding: 10px; text-align:   left; width: 60%">Descripcion</td>
                                <td style="padding: 10px; text-align: center; width: 10%">Expuesta</td>
                                <td style="padding: 10px; text-align: center; width: 15%">&nbsp;</td>
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
                                            <td style="padding: 5px; text-align: center; width: 15%;" onclick="editar('<%= Sistema %>', '<%= t("Parametro") %>')"><%= t("Parametro") %></td>
                                            <td style="padding: 5px; text-align:   left; width: 60%;" onclick="editar('<%= Sistema %>', '<%= t("Parametro") %>')"><%= t("Descripcion") %></td>
                                            <td style="padding: 5px; text-align: center; width: 10%;" onclick="editar('<%= Sistema %>', '<%= t("Parametro") %>')">
                                                <%
                                                    if t("Exponer") = 1 then   
                                                        response.write "Si"
                                                    else
                                                        response.write "&nbsp;"
                                                    end if
                                                %>
                                            </td>

                                            <td style="padding: 5px; text-align: right; width: 15%;">
                                                <button type="button" class="form-btn azul" onclick="asignar('<%= Sistema %>', '<%= t("Parametro") %>')">
                                                    <i class=" fa fa-user fa-xl" title="Asignar Roles"></i>
                                                </button>
                                                                                            
                                                <button type="button" class="form-btn rojo" onclick="borrar('<%= Sistema %>', '<%= t("Parametro") %>')">
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
                                response.write "No se encontró ninguna variable"
                            else
                                if conta = 1 then
                                    response.write "Se encontró una variable"
                                else
                                    response.write "Se encontraron " & conta & " variables"
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
                var sistema = document.getElementById("cboSistema").value;
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="variables.asp?s=" + sistema + "&o=" + ordenamiento;
                window.location.href = vinculo;                                      
            }
                  
            function editar(sistema, codigo) {
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var vinculo ="editar_variable.asp?s=" + sistema + "&p=" + codigo + "&o=" + ordenamiento;
                window.location.href = vinculo;
            }    

            function borrar(sistema, codigo) {
                var confirmacion = confirm("Desea borrar la variable seleccionada?");
                var vinculo ="borrar_variable.asp?s=" + sistema + "&p=" + codigo + "&o=<%= ordenadoPor %>";                

                if (confirmacion) {     
                    window.location.href = vinculo;
                };
            }    

            function asignar(sistema, parametro) {
                var vinculo ="/forma/desktop/apps/seguridad/usuarios/variables_asignar.asp?s=" + sistema + "&p=" + parametro + "&o=1";
                window.location.href = vinculo;
            }            
        </script> 

        <% cc.close: set cc = nothing %>      
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->                 
    </body>
</html>