<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Editar Roles</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0030"
            SysLockOut
        %>
        
        <style>
            .borde {
                border: 1px solid;
                border-color: rgb(184, 184, 184);
            }  
            
            .vbControl_B_Enabled {
                background-color: rgb(240, 255, 227);
                color: rgb(0, 0, 0);
               
                border: 1px solid rgb(28, 69, 117);
            }           
        </style>

        <%
            dim pCon, ptt, pProc, sqlString, Sistema, Proceso, ordenadoPor, tt
            dim Nombre, Descripcion, TipoRol, cuantosSistemas, primerTab

            Sistema = request.querystring("s")
            Proceso = request.querystring("p")
            ordenadoPor = request.querystring("o")

            set pCon = Server.CreateObject("ADODB.Connection")
            pCon.open Application("Conn")

            function NombreSistema(Sistema)
                dim sqlString

                sqlString = "SELECT sysNombre FROM seg_Sistemas WHERE (sysCodigo ='" & Sistema & "');"

                set ptt = pCon.execute(sqlString)   
                    NombreSistema = ptt("sysNombre") 
                ptt.close: set ptt = nothing
            end function   

            function CuantosUsuariosQuedan(Sistema, Proceso)
                sqlString = "SELECT COUNT(*) AS CUANTOS " & _
                             "FROM (SELECT TOP (100) PERCENT usuCodigo, usuNombre " & _
                                    "FROM dbo.seg_Usuarios " & _
                                   "WHERE (usuCodigo NOT IN (SELECT CodigoUsu " & _
                                                              "FROM dbo.seg_ProcesosUsuarios " & _
                                                             "WHERE (CodigoSis = '" & sistema & "') " & _
                                                               "AND (CodigoProc = '" & Proceso & "'))) " & _
                                     "AND (usuCodigo <> 'defaults') " & _
                                "ORDER BY usuNombre) AS t;"

                set ptt = pCon.execute(sqlString)   
                    CuantosUsuariosQuedan = ptt("Cuantos") 
                ptt.close: set ptt = nothing
            end function

            sub UsuariosAsignados(Sistema, Proceso)
                %>
                    <div class="tabla-wrapper">
                        <table class="tabla tabla-violet"> 
                            <thead>
                                <tr style="background-color: rgb(0, 0, 0); color: rgb(255, 255, 255);">
                                    <th class="sticky" style="width:  5%; text-align: center;">&nbsp;</th>
                                    <th class="sticky" style="width: 25%; text-align: center;  ">Usuario</th>
                                    <th class="sticky" style="width: 25%; text-align: center;  ">Cargo</th>
                                    <th class="sticky" style="width: 20%; text-align: center;  ">Tipo</th>
                                    <th class="sticky" style="width: 20%; text-align: center;  ">Estado</th>
                                    <th class="sticky" style="width:  5%; text-align: center;">&nbsp;</th>              
                                </tr>
                            </thead>

                            <tbody>
                                <% if CuantosUsuariosQuedan(Sistema, Proceso) > 0 then %>
                                    <tr>
                                        <td class="borde" style="text-align: center; font-size: 30px; font-weight: bold;">+</td>              

                                        <td colspan="4">
                                            <%
                                                sqlString = "SELECT usuCodigo, usuNombre " & _
                                                            "FROM dbo.seg_Usuarios " & _
                                                            "WHERE (usuCodigo NOT IN (SELECT CodigoUsu " & _
                                                                                        "FROM dbo.seg_ProcesosUsuarios " & _
                                                                                    "WHERE (CodigoSis = '" & Sistema & "') " & _
                                                                                        "AND (CodigoProc = '" & Proceso & "'))) " & _
                                                            "AND (usuCodigo <> 'defaults') " & _
                                                        "ORDER BY usuNombre;"

                                                set tt = pCon.execute(sqlString)
                                                    response.write "<select class='field full frame' name='usuNuevo' id='usuNuevo'>"
                                                        response.write "<option value='*'> - - Seleccione un Usuario - - </option>"

                                                        Do
                                                            response.write "<option value='" & tt("usuCodigo") & "'>" & tt("usuNombre") & "</option>"
                                                            tt.MoveNext
                                                        Loop Until tt.eof

                                                    response.write "</select>"
                                                tt.close: set tt = nothing
                                            %>                                    
                                        </td>              

                                        <td style="text-align: center;">
                                            <button type="button" class="form-btn verde" onclick="NuevaLinea('')">
                                                <i class="fa fa-save fa-xl"></i>
                                            </button>
                                        </td>              
                                    </tr>    
                                <% end if %>

                                <%
                                    sqlString = "SELECT p.CodigoSis, p.CodigoProc, p.CodigoUsu, p.Activo, p.TipoProceso, u.usuNombre, u.usuCargo " & _
                                                "FROM dbo.seg_ProcesosUsuarios AS p " & _
                                            "INNER JOIN dbo.seg_Usuarios AS u " & _
                                                    "ON p.CodigoUsu = u.usuCodigo " & _
                                                "WHERE (p.CodigoSis ='" & Sistema & "') " & _
                                                "AND (p.CodigoProc = '" & Proceso & "') " & _
                                            "ORDER BY u.usuNombre;"

                                    set cbox = pCon.execute(sqlString)
                                        if not (cbox.bof or cbox.eof) then
                                            contador = 0

                                            Do        
                                                contador = contador + 1

                                                response.write "<tr "
                                                    if cbox("Activo") <> "1" then
                                                        response.write "class='tr-rojo'"
                                                    end if                                        
                                                response.write ">"

                                                    response.write "<td style='text-align: center;'>"
                                                        response.write contador
                                                    response.write "</td>"      

                                                    response.write "<td style='text-align: left;'>"
                                                        response.write cbox("usuNombre")
                                                    response.write "</td>"                                   

                                                    response.write "<td style='text-align: left;'>"
                                                        response.write cbox("usuCargo")
                                                    response.write "</td>"

                                                    response.write "<td style='text-align: left;'>"
                                                        %>
                                                            <select name='pTipo<%= contador %>' 
                                                                    id='pTipo<%= contador %>' 
                                                                    class='field full frame' 
                                                                    onChange="ActualizarTipo('<%= cbox("CodigoUsu") %>','<%= cbox("Activo") %>', '<%= "pTipo" & contador %>')"
                                                                    style="<%
                                                                        if cbox("Activo") <> "1" then
                                                                            response.write "background-color: rgb(255, 222, 228);"
                                                                        end if                                                              
                                                                    %>">
                                                                <option value='1' <% if cbox("TipoProceso") = "1" then response.write " selected" %>>Permiso</option>
                                                                <option value='0' <% if cbox("TipoProceso") = "0" then response.write " selected" %>>Anti-Permiso</option>
                                                            </select>
                                                        <%
                                                    response.write "</td>" 

                                                    response.write "<td style='text-align: left;'>"
                                                        %>
                                                            <select name='pEstado<%= contador %>' 
                                                                    id='pEstado<%= contador %>' 
                                                                    class='field full frame' 
                                                                    onChange="ActualizarEstatus('<%= cbox("CodigoUsu") %>','<%= cbox("Activo") %>', '<%= "pEstado" & contador %>')"
                                                                    style="<%
                                                                        if cbox("Activo") <> "1" then
                                                                            response.write "background-color: rgb(255, 222, 228);"
                                                                        end if                                                              
                                                                    %>">
                                                                <option value='1' <% if cbox("Activo") = "1" then response.write " selected" %>>Activo</option>
                                                                <option value='0' <% if cbox("Activo") = "0" then response.write " selected" %>>Desactivado</option>
                                                            </select>
                                                        <%
                                                    response.write "</td>" 

                                                    response.write "<td style='text-align: center;'>"
                                                        %><button class="form-btn rojo" type="button" onClick="BorrarPermiso('<%= cbox("CodigoUsu") %>')"><%
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
                        </table> 
                    </div>
                <%     
            end sub                  
        %>
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->   

        <%
            set ptt = pCon.execute("SELECT * FROM seg_Procesos WHERE (proSistema = '" & Sistema & "') AND (proCodigo = '" & Proceso & "');")
                Nombre = ptt("proNombre") & " (" & ptt("proCodigo") & ")"
            ptt.close: set ptt = nothing
        %>

        <br />

        <div style="width: 95%; margin: auto;">
            <table style="width: 100%;">
                <tr>
                    <td style="width: 70%;">
                        <%
                            response.write "<span style='font-size: 22px'>" & Nombre & "</span><br />"
                        %>
                    </td>

                    <td style="width: 30%; text-align: right;">
                        <button class='form-btn rojo normal'  type='button' onclick="volver()">Volver</button>&nbsp;&nbsp;
                    </td>
                </tr>
            </table> 
        </div>

        <div style="width: 98%; margin: auto;">
            <form id="formulario"  name="formulario" method="post" action="grabar_rol.asp">
                <input id="ordenadoPor" name="ordenadoPor" type="text" value="<%= ordenadoPor %>" class="no-ver"/>
                <input id="Codigo" name="Codigo" type="text" value="<%= Rol %>" class="no-ver"/>

                <div class="main main-scroll">
                    <div class="line label-top">
                        <label class="label normal">Permisos</label>
                        <label class="label full section">
                            <% UsuariosAsignados Sistema, Proceso %> 
                        </label>
                    </div>
                </div>
            </form>
        </div>

        <br /><br />

        <script>    
            function volver() {
                var vinculo = "permisos.asp?s=<%= Sistema %>&o=<%= ordenadoPor %>";
                window.location.href = vinculo;
            }

            function NuevaLinea() {
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var usuNuevo = document.getElementById("usuNuevo").value;

                if (usuNuevo == "*") {
                    window.alert("Debe Seleccionar un Usuario");
                } else {
                    var vinculo ="permisos_asignar_usuario.asp?o=" + ordenamiento + "&s=<%= Sistema %>&p=<%= Proceso %>&u=" + usuNuevo;
                    window.location.href = vinculo;    
                }
            }

            function ActualizarEstatus(usuario, estactual, comboBox) {
                var cboEstatus = document.getElementById(comboBox);
                var confirmacion = confirm("Desea actualizar el Estado del usuario seleccionado?");
                var vinculo ="permisos_estatus.asp?s=<%= Sistema %>&p=<%= Proceso %>&u=" + usuario + "&e=" + cboEstatus.value +  "&o=<%= ordenadoPor %>";     
                
                if (confirmacion) {     
                    window.location.href = vinculo;
                }
                else {
                    cboEstatus.value = estactual;
                };                
            }

            function ActualizarTipo(usuario, estactual, comboBox) {
                var cboEstatus = document.getElementById(comboBox);
                var confirmacion = confirm("Desea actualizar el Tipo de Permiso del usuario seleccionado?");
                var vinculo ="permisos_tipo.asp?s=<%= Sistema %>&p=<%= Proceso %>&u=" + usuario + "&e=" + cboEstatus.value +  "&o=<%= ordenadoPor %>";     
                
                if (confirmacion) {     
                    window.location.href = vinculo;
                }
                else {
                    cboEstatus.value = estactual;
                };                
            }            

            function BorrarPermiso(usuario) {
                var confirmacion = confirm("Desea quitar de este permiso al usuario seleccionado?");
                var vinculo ="permisos_asignar_borrar.asp?o=<%= ordenadoPor %>&s=<%= Sistema %>&p=<%= Proceso %>&u=" + usuario;

                if (confirmacion) {     
                    window.location.href = vinculo;
                };
            }      
        </script> 

        <% pCon.close: set pCon = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->   
    </body>
</html>