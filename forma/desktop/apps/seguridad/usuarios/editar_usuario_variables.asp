<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Asignar Variables</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0040"
            SysLockOut
        %>   
        
        <style>
            .borde {
                border: 1px solid;
                border-color: rgb(184, 184, 184);
            }  

            select option { 
                line-height: 2.0;
            }            
        </style>

        <%
            dim pCon, ptt, pProc, sqlString, Usuario, ordenadoPor, tt
            dim Nombre, Descripcion, TipoRol

            Usuario = request.querystring("u")
            ordenadoPor = request.querystring("o")

            set pCon = Server.CreateObject("ADODB.Connection")
            pCon.open Application("Conn")

            function NombreUsuario(Usuario)
                sqlString = "SELECT usuNombre FROM seg_Usuarios " & _
                             "WHERE (usuCodigo = '" & Usuario & "');"

                set ptt = pCon.execute(sqlString)   
                    NombreUsuario = ptt("usuNombre") 
                ptt.close: set ptt = nothing
            end function            

            Function ConsultaSQL(Sistema, Parametro) 
                dim sqlString

                sqlString = "SELECT ConsultaSQL AS Consulta FROM seg_Parametros WHERE (Sistema = '" & Sistema & "' AND Parametro = '" & Parametro & "');"

                set ptt = pCon.execute(sqlString)   
                    ConsultaSQL = ptt("Consulta") 
                    ConsultaSQL = Replace(ConsultaSQL, "|", "'")
                ptt.close: set ptt = nothing                
            End Function            

            function ParamNoAsignados(Usuario)
                sqlString = "SELECT COUNT(*) AS Cuantos FROM (" & _
                            "SELECT p.Sistema + '__' + p.Parametro AS Variable, p.Sistema, p.Parametro, s.sysNombre + ':  ' + p.Descripcion AS nParam " & _
                            "FROM dbo.seg_Parametros AS p INNER JOIN dbo.seg_Sistemas AS s ON p.Sistema = s.sysCodigo INNER JOIN (" & _
                            "SELECT DISTINCT Sistema FROM dbo.seg_PermisosUsuarios WHERE (Usuario = '" & Usuario & "')) AS pu ON s.sysCodigo = pu.Sistema " & _
                            "WHERE ((p.Sistema + '__' + p.Parametro) NOT IN (SELECT Sistema + '__' + Parametro AS V FROM dbo.seg_Usuarios_Parametros " & _
                            "WHERE (Usuario = '" & Usuario & "'))) " & _
                            ") AS q;"
                set ptt = pCon.execute(sqlString)   
                    ParamNoAsignados = ptt("Cuantos") 
                ptt.close: set ptt = nothing
            end function

            sub VariablesAsignadas(Usuario)      
                %>
                    <div class="tabla-wrapper">
                        <table class="tabla tabla-green"> 
                            <thead>
                                <tr>
                                    <th class="sticky" style="width:  5%; text-align: center;">&nbsp;</td>
                                    <th class="sticky" style="width: 15%; text-align: left;  ">Sistema</td>
                                    <th class="sticky" style="width: 15%; text-align: left;  ">Variable</td>
                                    <th class="sticky" style="width: 40%; text-align: left;  ">Descripcion</td>
                                    <th class="sticky" style="width: 20%; text-align: left;  ">Valor</td>
                                    <th class="sticky" style="width:  5%; text-align: center;">&nbsp;</td>              
                                </tr>
                            </thead>

                            <tbody>
                                <% if ParamNoAsignados(Usuario) > 0 then %>
                                    <tr>
                                        <td>+</td>

                                        <td colspan="4">
                                            <%
                                                sqlString = "SELECT p.Sistema + '__' + p.Parametro AS Variable, p.Sistema, p.Parametro, s.sysNombre + ':  ' + p.Descripcion AS nParam " & _
                                                            "FROM dbo.seg_Parametros AS p INNER JOIN dbo.seg_Sistemas AS s ON p.Sistema = s.sysCodigo INNER JOIN (" & _
                                                            "SELECT DISTINCT Sistema FROM dbo.seg_PermisosUsuarios WHERE (Usuario = '" & Usuario & "')) AS pu ON s.sysCodigo = pu.Sistema " & _
                                                            "WHERE ((p.Sistema + '__' + p.Parametro) NOT IN (SELECT Sistema + '__' + Parametro AS V FROM dbo.seg_Usuarios_Parametros " & _
                                                            "WHERE (Usuario = '" & Usuario & "')));"                                            

                                                set tt = pCon.execute(sqlString)
                                                    response.write "<select class='field frame full' name='varNueva' id='varNueva'>"
                                                        response.write "<option value='*'> - - Seleccione una Variable - - </option>"

                                                        Do
                                                            response.write "<option value='" & tt("Variable") & "'>" & tt("nParam") & "</option>"
                                                            tt.MoveNext
                                                        Loop Until tt.eof

                                                    response.write "</select>"
                                                tt.close: set tt = nothing
                                            %>                                    
                                        </td>              

                                        <td>
                                            <button type="button" class="form-btn verde" onclick="NuevaLinea()"><i class="fa fa-save fa-xl"></i></button>
                                        </td>              
                                    </tr>    
                                <% end if %>

                                <%
                                    sqlString = "SELECT up.Sistema, up.Parametro, p.TipoParametro, s.sysNombre, p.Descripcion, up.Valor, up.Sistema + '__' + up.Parametro AS Variable " & _
                                                "FROM dbo.seg_Parametros AS p INNER JOIN dbo.seg_Sistemas AS s ON p.Sistema = s.sysCodigo " & _
                                                "INNER JOIN dbo.seg_Usuarios_Parametros AS up ON p.Parametro = up.Parametro AND p.Sistema = up.Sistema " & _
                                                "WHERE (up.Usuario = '" & Usuario & "') ORDER BY s.sysNombre, p.Parametro;"

                                    set cbox = pCon.execute(sqlString)
                                        if not (cbox.bof or cbox.eof) then
                                            contador = 0

                                            Do        
                                                contador = contador + 1
                                                nomCampo = "pValor_" & contador

                                                response.write "<tr>"
                                                    response.write "<td style='text-align: center;'>"
                                                        response.write contador
                                                    response.write "</td>"      

                                                    response.write "<td style='text-align: left;'>"
                                                        response.write cbox("sysNombre")
                                                    response.write "</td>"                                             

                                                    response.write "<td style='text-align: left;'>"
                                                        response.write cbox("Parametro")
                                                    response.write "</td>"                                   

                                                    response.write "<td style='text-align: left;'>"
                                                        response.write cbox("Descripcion")
                                                    response.write "</td>"

                                                    response.write "<td style='text-align: left;'>"
                                                        Select Case cbox("TipoParametro")
                                                            Case "1"
                                                                '
                                                                ' Es un Permiso... El Campo tiene un valor de 1 y está escondido
                                                                '
                                                                response.write "<input class='field full frame' id='" & nomCampo & " name=" & nomCampo & " type='text' "
                                                                response.write "value=" & cbox("Valor") & " style='visibility: hidden;' />"
                                                            Case "2"
                                                                '
                                                                ' El Campo es una Variable... Presentamos el input-box
                                                                '
                                                                %>
                                                                    <input class="field full frame" id="<%= nomCampo %>" name="<%= nomCampo %>" type="text" value="<%= cbox("Valor") %>" 
                                                                        onChange="actualizarValor('<%= nomCampo %>', '<%= cbox("Variable") %>', '<%= cbox("Valor") %>', 2)" />
                                                                <%
                                                            Case "3"
                                                                '
                                                                ' La Variable es un Campo "Si/No". Presentamos un Check-Box
                                                                '
                                                                %>
                                                                    <input type="radio" id="<%= nomCampo & "1" %>" name="<%= nomCampo %>" value="1" <%
                                                                        if cbox("Valor") = "1" then response.write " checked " %>
                                                                        onChange="actualizarValor('<%= nomCampo & "1" %>', '<%= cbox("Variable") %>', '<%= cbox("Valor") %>', 3)">
                                                                    &nbsp;&nbsp;
                                                                    <label style="font-family: Verdana; font-size: 14px;" for="Si">Si</label>

                                                                    &nbsp;&nbsp;&nbsp;&nbsp;

                                                                    <input type="radio" id="<%= nomCampo & "0" %>" name="<%= nomCampo %>" value="0" <%
                                                                        if cbox("Valor") = "0" then response.write " checked " %>
                                                                        onChange="actualizarValor('<%= nomCampo & "0" %>', '<%= cbox("Variable") %>', '<%= cbox("Valor") %>', 3)">
                                                                    &nbsp;&nbsp;
                                                                    <label style="font-family: Verdana; font-size: 14px;" for="No">No</label>
                                                                <%

                                                            Case "4"
                                                                '
                                                                ' Es una lista... Presentamos un Combo-Box
                                                                '
                                                                set lista = pCon.execute("SELECT * FROM seg_Parametros_Valores WHERE Sistema = '" & cbox("Sistema") & "' AND Parametro = '" & cbox("Parametro") & "' ORDER BY Descripcion;")

                                                                if not (lista.bof or lista.eof) then
                                                                    %> <select name="<%= nomCampo %>" id="<%= nomCampo %>" class="field full frame" onChange="actualizarValor('<%= nomCampo %>', '<%= cbox("Variable") %>', '<%= cbox("Valor") %>', 4)"> <%
                                                                        do
                                                                            response.write "<option value='" & lista("Valor") & "'"
                                                                                if cbox("Valor") = lista("Valor") then
                                                                                    response.write " selected"
                                                                                end if
                                                                            response.write ">" & lista("Descripcion") & "</option>"
                                                                            
                                                                            lista.MoveNext
                                                                        loop until (lista.eof)
                                                                    response.write "<select>"
                                                                end if

                                                                lista.close: set lista = nothing

                                                            Case "5"
                                                                '
                                                                ' Es una Consulta... Presentamos un Combo-Box
                                                                '
                                                                Consulta = ConsultaSQL(cbox("Sistema"), cbox("Parametro"))

                                                                set lista = pCon.execute(Consulta)

                                                                if not (lista.bof or lista.eof) then
                                                                    %><select name="<%= nomCampo %>" id="<%= nomCampo %>" class="field full frame" onChange="actualizarValor('<%= nomCampo %>', '<%= cbox("Variable") %>', '<%= cbox("Valor") %>', 5)"> <%
                                                                        do
                                                                            response.write "<option value='" & lista("Codigo") & "'"
                                                                                if cbox("Valor") = lista("Codigo") then
                                                                                    response.write " selected"
                                                                                end if
                                                                            response.write ">" & lista("Valor") & "</option>"
                                                                            
                                                                            lista.MoveNext
                                                                        loop until (lista.eof)
                                                                    response.write "<select>"
                                                                end if

                                                                lista.close: set lista = nothing    

                                                            Case "6"
                                                                '
                                                                ' El Campo es un selector de color... Presentamos la rueda de colores
                                                                '
                                                                %>
                                                                    <input class="field full frame" 
                                                                        name="<%= nomCampo %>" 
                                                                        id="<%= nomCampo %>" 
                                                                        type="color" value="<%= cbox("Valor") %>"
                                                                        onChange="actualizarValor('<%= nomCampo %>', '<%= cbox("Variable") %>', '<%= cbox("Valor") %>', 6)">
                                                                <%

                                                            Case 7
                                                                '
                                                                ' El Campo es una Barra de Desplazamiento
                                                                '
                                                                %>

                                                                <label style="width: 100%;">
                                                                    <input class="field full" style="width: 100%" type="range" min="0" max="100" 
                                                                        value="<%= cbox("Valor") %>" 
                                                                        id="<%= nomCampo %>" name="<%= nomCampo %>"
                                                                        onChange="actualizarValor('<%= nomCampo %>', '<%= cbox("Variable") %>', '<%= cbox("Valor") %>', 7)">
                                                                </label>
                                                            <%                                                       
                                                        End Select
                                                    response.write "</td>" 

                                                    response.write "<td style='text-align: center; width: 5%' class='borde'>"
                                                        %><button class="form-btn rojo" type="button" onClick="BorrarVariable('<%= cbox("Variable") %>')"><%
                                                            response.write "<i class=' fa fa-trash fa-xl' title='Borrar Detalle'></i>"
                                                        response.write "</button>"                    
                                                    response.write "</td>"                
                                                response.write "</tr>"

                                                cbox.MoveNext
                                            Loop Until (cbox.eof)
                                        end if
                                    cbox.close: set cbox = nothing     
                                %>
                            </tbody>    

                            <tfoot>
                                <tr>
                                    <td class="sticky" style="text-align: center;" colspan="6">
                                        <%
                                            Select Case contador
                                                case 0: response.write "No se encontraron variables"
                                                case 1: response.write "Se encontró una variable"
                                                case else
                                                    response.write "Se encontraron " & contador & " variables"
                                            end select
                                        %>
                                    </td>
                                </tr>
                            </tfoot>               
                        </table> 
                    </div>
                <%     
            end sub                  
        %>
    </head>

    <body plantilla="tabla" reserva="190">
        <!-- #include virtual = "/core/includes/kernel/body.inc" --> 

        <br />

        <div style="width: 95%; margin: auto;">
            <table style="width: 100%;">
                <tr>
                    <td style="width: 60%; font-size: 22px;">
                        <h3></i>Asignar Variables a <%= NombreUsuario(usuario) %></h3>
                    </td>

                    <td style="width: 40%; text-align: right;">
                        <button class='form-btn rojo normal'  type='button' onclick="volver()">Volver</button>&nbsp;&nbsp;
                    </td>
                </tr>
            </table>
        </div>

        <div style="width: 98%; margin: auto;">
            <div class="main">
                <div class="line">
                    <div class="label full section">
                        <% VariablesAsignadas Usuario %> 
                    </div>
                </div>
            </div>
        </div>

        <script>
            function volver() {
                var vinculo = "lista.asp?o=<%= ordenadoPor %>";
                window.location.href = vinculo;
            }

            function NuevaLinea() {
                var varNueva = document.getElementById("varNueva").value;

                if (varNueva == "*") {
                    window.alert("Debe Seleccionar una Variable");
                } else {
                    var vinculo ="variables_asignar_usuario_l.asp?o=<%= ordenadoPor %>&p=" + varNueva + "&u=<%= Usuario %>";
                    window.location.href = vinculo;    
                }
            }

            function actualizarValor(campo, parametro, valorActual, tipoCampo) {
                var txtCampo = document.getElementById(campo);
                var nuevo_valor = document.getElementById(campo).value;

                if (tipoCampo == "6") {
                    nuevo_valor = nuevo_valor.replace("#", "^");
                };

                var confirmacion = confirm("Desea actualizar el valor de la variable?");
                var vinculo ="variables_actualizar_l.asp?p=" + parametro + "&u=<%= Usuario %>&v=" + nuevo_valor +  "&o=<%= ordenadoPor %>" + "&t=" + tipoCampo;     

                if (confirmacion) {     
                    window.location.href = vinculo;
                }
                else {
                    txtCampo.value = valorActual;
                };                
            }

            function BorrarVariable(parametro) {
                var confirmacion = confirm("Desea quitar la variable " + parametro + "?");
                var vinculo = "variables_asignar_borrar_l.asp?o=<%= ordenadoPor %>&p=" + parametro + "&u=<%= Usuario %>";

                if (confirmacion) {     
                    window.location.href = vinculo;
                };
            }      
        </script> 

        <% pCon.close: set pCon = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>