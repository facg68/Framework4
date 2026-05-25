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
            thisProcess = "seg.0025"
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
                padding: 5px;
                border: 1px solid rgb(28, 69, 117);
            }  
        </style>

        <%
            dim pCon, ptt, pProc, sqlString, Rol, ordenadoPor, tt, Unico, vinculo
            dim Nombre, Descripcion, TipoRol, cuantosSistemas, primerTab, Sistema

            Rol = request.querystring("r")
            Sistema = request.querystring("s")
            ordenadoPor = request.querystring("o")
            Unico = request.querystring("un")

            set pCon = Server.CreateObject("ADODB.Connection")
            pCon.open Application("Conn")

            function NombreSistema(Sistema)
                dim sqlString

                sqlString = "SELECT sysNombre FROM seg_Sistemas WHERE (sysCodigo ='" & Sistema & "');"

                set ptt = pCon.execute(sqlString)   
                    NombreSistema = ptt("sysNombre") 
                ptt.close: set ptt = nothing
            end function   

            function CuantosUsuariosQuedan(Rol)
                sqlString = "SELECT COUNT(*) AS Cuantos " & _
                            "FROM (SELECT usuCodigo, usuNombre " & _
                                   "FROM dbo.seg_Usuarios " & _
                                   "WHERE (usuCodigo NOT IN (" & _
                                           "SELECT CodigoUsuario " & _
                                           "FROM dbo.seg_RolesUsuarios " & _
                                           "WHERE (CodigoRol = '" & Rol & "'))) " & _
                                   "AND (usuCodigo <> 'defaults')) AS t;"

                set ptt = pCon.execute(sqlString)   
                    CuantosUsuariosQuedan = ptt("Cuantos") 
                ptt.close: set ptt = nothing
            end function

            sub UsuariosAsignados(Rol)
                %>
                    <table style="width: 100%;" style="padding: 0px;"> 
                        <tr style="background-color: rgb(0, 0, 0); color: rgb(255, 255, 255);">
                            <td class="top borde" style="width:  5%; text-align: center; padding: 5px;">&nbsp;</td>
                            <td class="top borde" style="width: 45%; text-align: left;   padding: 5px;">Usuario</td>
                            <td class="top borde" style="width: 35%; text-align: left;   padding: 5px;">Cargo</td>
                            <td class="top borde" style="width: 10%; text-align: left;   padding: 5px;">Estado</td>
                            <td class="top borde" style="width:  5%; text-align: center; padding: 5px;">&nbsp;</td>              
                        </tr>

                        <% if CuantosUsuariosQuedan(Rol) > 0 then %>
                            <tr>
                                <td class="borde" style="width: 5%; text-align: center; font-size: 30px; font-weight: bold;">+</td>              

                                <td colspan="3" class="borde" style="width: 55%;">
                                    <%
                                        sqlString = "SELECT usuCodigo, usuNombre " & _
                                                    "FROM dbo.seg_Usuarios " & _
                                                    "WHERE (usuCodigo NOT IN (" & _
                                                            "SELECT CodigoUsuario " & _
                                                            "FROM dbo.seg_RolesUsuarios " & _
                                                            "WHERE (CodigoRol = '" & Rol & "'))) " & _
                                                    "AND (usuCodigo <> 'defaults') " & _
                                                    "ORDER BY usuNombre;"

                                        set tt = pCon.execute(sqlString)
                                            response.write "<select class='field full' name='usuNuevo' id='usuNuevo'>"
                                                response.write "<option value='*'> - - Seleccione un Usuario - - </option>"

                                                Do
                                                    response.write "<option value='" & tt("usuCodigo") & "'>" & tt("usuNombre") & "</option>"
                                                    tt.MoveNext
                                                Loop Until tt.eof

                                            response.write "/<select>"
                                        tt.close: set tt = nothing
                                    %>                                    
                                </td>              

                                <td class="borde" style="width: 5%; text-align: center; background-color: rgb(240, 255, 227);">
                                    <button type="button" class="form-btn verde" onclick="NuevaLinea('')"><i class="fa fa-save fa-xl"></i></button>
                                </td>              
                            </tr>    
                        <% end if %>

                        <%
                            sqlString = "SELECT ru.CodigoRol, ru.CodigoUsuario, u.usuNombre, u.usuCargo, ru.Activo " & _
                                        "FROM dbo.seg_RolesUsuarios AS ru " & _
                                        "INNER JOIN dbo.seg_Usuarios AS u " & _
                                        "ON ru.CodigoUsuario = u.usuCodigo " & _
                                        "WHERE (ru.CodigoRol = '" & Rol & "') " & _
                                        "ORDER BY u.usuNombre;"

                            set cbox = pCon.execute(sqlString)
                                if not (cbox.bof or cbox.eof) then
                                    contador = 0

                                    Do        
                                        contador = contador + 1

                                        response.write "<tr style='"
                                            if cbox("Activo") = "1" then
                                                response.write "background-color: rgb(255,255,255);"
                                            else
                                                response.write "background-color: rgb(255, 222, 228);"
                                            end if                                        
                                        response.write "'>"

                                            response.write "<td style='text-align: center; padding: 5px; width: 5%' class='borde'>"
                                                response.write contador
                                            response.write "</td>"      

                                            response.write "<td style='text-align: left; padding: 5px; width: 45%' class='borde'>"
                                                response.write cbox("usuNombre")
                                            response.write "</td>"                                   

                                            response.write "<td style='text-align: left; padding: 5px; width: 35%' class='borde'>"
                                                response.write cbox("usuCargo")
                                            response.write "</td>"   

                                            response.write "<td style='text-align: left; padding: 5px; width: 35%' class='borde'>"
                                                %>
                                                    <select name='uEstado<%= contador %>' 
                                                            id='uEstado<%= contador %>' 
                                                            class='field full' 
                                                            onChange="ActualizarEstatus('<%= cbox("CodigoUsuario") %>','<%= cbox("Activo") %>', '<%= "uEstado" & contador %>')"
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

                                            response.write "<td style='text-align: center; padding: 5px; width: 5%' class='borde'>"
                                                %><button class="form-btn rojo" type="button" onClick="BorrarUsuRol('<%= cbox("CodigoUsuario") %>')"><%
                                                    response.write "<i class=' fa fa-trash fa-xl' title='Borrar Detalle'></i>"
                                                response.write "</button>"                    
                                            response.write "</td>"                
                                        response.write "</tr>"

                                        cbox.MoveNext
                                    Loop Until (cbox.eof)
                                end if
                            cbox.close: set cbox = nothing     
                        %>                   
                    </table> 
                <%     
            end sub                  
        %>
    </head>

    <body style="background-color: rgb(230, 230, 230);">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->   

        <%
            set ptt = pCon.execute("SELECT * FROM seg_Roles WHERE (rolCodigo = '" & Rol & "');")
                Nombre = ptt("rolNombre")
                Descripcion = ptt("rolDescripcion")
                TipoRol = ptt("TipoRol")
            ptt.close: set ptt = nothing
        %>

        <div style="width: 95%; margin: auto;">
            <br />

            <table style="width: 100%;">
                <tr>
                    <td style="width: 70%; font-size: 24px;">
                        <%
                            response.write "<span style='font-size: 22px'>" & Nombre & "&nbsp;("
                                if TipoRol = 1 then
                                    response.write "Rol"
                                else
                                    response.write "Anti-Rol"
                                end if
                            response.write ")</span><br />"

                            response.write "<span style='font-size: 16px'>" & Descripcion & "</span>"
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
                <input id="ordenadoPor" name="ordenadoPor"  type="text" value="<%= ordenadoPor %>"  class="no-ver"/>
                <input id="Codigo"      name="Codigo"       type="text" value="<%= Rol %>"          class="no-ver"/>
                <input id="Unico"       name="Unico"        type="text" value="<%= Unico %>"        class="no-ver"/>
                <input id="Sistema"     name="Sistema"      type="text" value="<%= Sistema %>"      class="no-ver"/>

                <div class="main main-scroll">
                    <div class="line">
                        <label class="label normal"><%
                            if TipoRol = 1 then
                                response.write "Se Permite"
                            else
                                response.write "Se Eliminan"
                            end if                                                
                        %></label>

                        <label class="label full section">
                            <%
                                sqlString = "SELECT detRolSistema AS sysCodigo, sysNombre, Procesos " & _
                                            "FROM (" & _
                                                    "SELECT rd.detRolSistema, s.sysNombre, COUNT(rd.detRolProceso) AS Procesos " & _
                                                    "FROM dbo.seg_RolDetalles AS rd INNER JOIN dbo.seg_Sistemas AS s ON rd.detRolSistema = s.sysCodigo " & _
                                                    "WHERE (rd.detRol = '" & Rol & "') GROUP BY rd.detRolSistema, s.sysNombre " & _
                                                ") AS t " & _
                                            "WHERE (Procesos > 1) " & _
                                            "ORDER BY sysNombre;"

                                set ptt = pCon.execute(sqlString)

                                if not (ptt.bof or ptt.eof) then
                                    response.write "<table style='width: 100%; border: none; border-spacing: 0px;'>"
                                        response.write "<tr>"
                                            cuantosSistemas = 0
                                            primerTab = "/*"

                                            do
                                                cuantosSistemas = cuantosSistemas + 1
                                                if primerTab = "/*" then primerTab = ptt("sysCodigo")
                                                %>
                                                    <td style="padding: 10px; border: 1px solid rgb(187, 188, 189); 
                                                                text-align: center; background-color: rgb(200, 202, 204);"
                                                        onclick="Tabs_Display('<%= ptt("sysCodigo") %>')" >
                                                        <%= ptt("sysNombre") %>
                                                    </td>
                                                <%

                                                ptt.movenext
                                            loop until ptt.eof
                                        response.write "</tr>"   

                                        ptt.MoveFirst

                                        response.write "<tr>"
                                            response.write "<td colspan='" & cuantosSistemas & "' " 
                                            response.write "style='padding: 10px; border: 1px solid rgb(187, 188, 189); text-align: center; background-color: rgb(255, 255, 255);'>"

                                                do
                                                    sqlString = "SELECT rd.detRolSistema, p.proNombre " & _
                                                                "FROM dbo.seg_RolDetalles AS rd INNER JOIN dbo.seg_Procesos AS p " & _
                                                                "ON rd.detRolProceso = p.proCodigo AND rd.detRolSistema = p.proSistema " & _
                                                                "WHERE (rd.detRol = '" & Rol & "') AND (rd.detRolSistema = '" & ptt("sysCodigo") & "') " & _
                                                                "ORDER BY p.proNombre ASC;"

                                                    response.write "<div id = '" & ptt("sysCodigo") & "' style='display: "
                                                        if ptt("sysCodigo") = primerTab then 
                                                            response.write "block"
                                                        else
                                                            response.write "none"
                                                        end if
                                                    response.write "; text-align: left; font-size: 12px; line-height:1.8em;'>"
                                                        set pProc = pCon.execute(sqlString)

                                                        if not (pProc.bof or pProc.eof) then
                                                            do
                                                                response.write pProc("proNombre") & "<br />"
                                                                pProc.movenext
                                                            loop until pProc.eof
                                                        end if

                                                        pProc.close: set pProc = nothing                                
                                                    response.write "</div>"

                                                    ptt.MoveNext
                                                loop until ptt.eof

                                            response.write "</td>"
                                        response.write "</tr>"

                                    response.write "</table>"

                                    ptt.close: set ptt = nothing
                                end if
                            %>                                          
                        </label>
                    </div>

                    <div class="line">
                        <label class="label normal">Usuarios</label>
                        <label class="label full section">
                            <% UsuariosAsignados Rol %> 
                        </label>
                    </div>
                </div>
            </form>
        </div>

        <script>
            pageReserva = 165;

            function volver() {
                var vinculo ="<%
                    if unico = "1" then
                        response.write "../sistemas/roles_sys.asp?s=" & Sistema & "&o=" & OrdenadoPor
                    else
                        response.write "../sistemas/roles.asp?o=" & OrdenadoPor
                    end if                
                %>";

                window.location.href = vinculo;
            }

            function NuevaLinea() {
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var usuNuevo = document.getElementById("usuNuevo").value;

                if (usuNuevo == "*") {
                    window.alert("Debe Seleccionar un Usuario");
                } else {
                    var vinculo ="roles_asignar_usuario.asp?o=" + ordenamiento + "&r=<%= Rol %>&u=" + usuNuevo + "&un=<%= Unico %>" + "&s=<%= Sistema %>";
                    window.location.href = vinculo;
                }
            }

            function ActualizarEstatus(usuario, estactual, comboBox) {
                var cboEstatus = document.getElementById(comboBox);
                var confirmacion = confirm("Desea actualizar el Estado del usuario seleccionado?");
                var vinculo ="roles_estatus.asp?r=<%= Rol %>&u=" + usuario + "&e=" + estactual +  "&o=<%= ordenadoPor %>" + "&un=<%= Unico %>"+ "&s=<%= Sistema %>";     
                
                if (confirmacion) {     
                    window.location.href = vinculo;
                }
                else {
                    cboEstatus.value = estactual;
                };                
            }

            function BorrarUsuRol(usuario) {
                var confirmacion = confirm("Desea quitar de este rol al usuario seleccionado?");
                var vinculo ="roles_asignar_borrar.asp?r=<%= Rol %>&u=" + usuario + "&o=<%= ordenadoPor %>" + "&un=<%= Unico %>"+ "&s=<%= Sistema %>";                

                if (confirmacion) {     
                    window.location.href = vinculo;
                };
            }

            function Tabs_Display(codigo) {
                var e = document.getElementById(codigo);

                Tabs_Reset();
                e.style.display =  'block';
            }

            function Tabs_Reset() {
                <%
                    sqlString = "SELECT sysCodigo, sysNombre FROM dbo.seg_Sistemas AS s ORDER BY sysNombre;"
                    set ptt = pCon.execute(sqlString)

                    cuantosSistemas = 0

                    if not (ptt.bof or ptt.eof) then
                        do
                            cuantosSistemas = cuantosSistemas + 1
                            response.write "var e" & cuantosSistemas & " = document.getElementById('" & ptt("sysCodigo") & "');"

                            ptt.movenext
                        loop until ptt.eof

                        for ckk = 1 to cuantosSistemas
                            response.write "e" & ckk & ".style.display = 'none';"
                        next
                    end if

                    ptt.close: set ptt = nothing                
                %>
            }     
        </script> 

        <% pCon.close: set pCon = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->   
    </body>
</html>