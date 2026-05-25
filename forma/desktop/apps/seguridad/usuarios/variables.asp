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
            thisProcess = "seg.0040"
            SysLockOut
        %>                        

        <%
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
            if ordenadoPor = "" then ordenadoPor = "0"

            sqlString = "SELECT Parametro, Descripcion FROM seg_Parametros WHERE (Sistema = '" & sistema & "') "

            select case ordenadoPor
                case 0: sqlString = sqlString & " ORDER BY Parametro;"
                case 1: sqlString = sqlString & " ORDER BY Parametro DESC;"
            end select

            set t = cc.execute(sqlString)        
        %>

        <div style="width: 98%; margin: auto;">
            <table style="width:100%; margin: auto;">
                <tr style="padding: 10px;">
                    <td>
                        <table style="width: 100%;">                
                            <tr>
                                <td style="text-align:left; width: 60%">
                                    <span style="font-size: 24px">&nbsp;Asignar Variables  - <%= NombreSistema(Sistema) %></span>

                                    <br />

                                    <span style="font-size: 20px">
                                        &nbsp;Ordenados por&nbsp;
                                        <%
                                            select case ordenadoPor
                                                case 0: response.write "Nombre"
                                                case 1: response.write "Nombre (descendentemente)"
                                            end select                            
                                        %>
                                    </span>                            
                                </td>

                                <td style="text-align:right; width: 40%">
                                    <select name="cboSistema" id="cboSistema" required 
                                            Class="field"
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
                                        <option value="0" <% if ordenadoPor = "0" then response.write " selected" %>>&#9650; Nombre</option>
                                        <option value="1" <% if ordenadoPor = "1" then response.write " selected" %>>&#9660; Nombre</option>
                                    </select>         
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
                                <td style="padding: 10px; text-align:center; width: 15%">Nombre</td>
                                <td style="padding: 10px; text-align:  left; width: 85%">Descripcion</td>
                                <td style="padding: 10px; text-align:center; width:  5%">&nbsp;</td>
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
                                            <td style="padding: 5px; text-align:center; width: 15%;" onclick="asignar('<%= Sistema %>', '<%= t("Parametro") %>')"><%= t("Parametro") %></td>
                                            <td style="padding: 5px; text-align:  left; width: 80%;" onclick="asignar('<%= Sistema %>', '<%= t("Parametro") %>')"><%= t("Descripcion") %></td>

                                            <td style="padding: 5px; text-align: right; width: 5%;">
                                                <button type="button" class="form-btn azul" onclick="asignar('<%= Sistema %>', '<%= t("Parametro") %>')">
                                                    <i class=" fa fa-user fa-xl" title="Asignar Variable"></i>
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

        <%
            t.close: set t = nothing
        %>

        <script>
            function filtrar() {
                var sistema = document.getElementById("cboSistema").value;
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="variables.asp?s=" + sistema + "&o=" + ordenamiento;
                window.location.href = vinculo;                                      
            }
                  
            function asignar(sistema, codigo) {
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var vinculo ="variables_asignar.asp?s=" + sistema + "&p=" + codigo + "&o=" + ordenamiento;
                window.location.href = vinculo;
            }          
        </script> 

        <% cc.close: set cc = nothing %>
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>