<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Editar Usuario - Asignar Roles</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->        
        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0025"
            SysLockOut

            dim cc, tt, sqlString, usuario, ordenadoPor

            usuario = Request.querystring("u")
            ordenadoPor = Request.querystring("o")

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")

            '
            ' Funciones y Procedimientos
            '

            function NombreUsuario(usuario)     
                set tt = cc.execute("SELECT usuNombre FROM seg_Usuarios WHERE usuCodigo ='" & usuario & "';")
                    NombreUsuario = tt("usuNombre")
                tt.close: set tt = nothing
            end function

            function procRoles(Rol)
                dim sqlCommand

                procRoles = ""
                sqlCommand = "SELECT detRol, Proceso " & _
                                "FROM (SELECT rd.detRol, s.sysNombre + '  |  ' + p.proNombre AS Proceso " & _
                                        "FROM dbo.seg_RolDetalles AS rd " & _
                                    "INNER JOIN dbo.seg_Procesos AS p " & _
                                            "ON rd.detRolProceso = p.proCodigo " & _
                                        "AND rd.detRolSistema = p.proSistema " & _
                                    "INNER JOIN dbo.seg_Sistemas AS s " & _
                                            "ON p.proSistema = s.sysCodigo " & _
                                        "WHERE (rd.detRol = '" & Rol & "')) AS r " & _
                            "ORDER BY Proceso;"

                set pp = cc.execute(sqlCommand)
                    if not (pp.bof or pp.eof) then
                        do
                            procRoles = procRoles & pp("Proceso") & "&#10;"   
                            pp.MoveNext
                        loop until pp.eof
                    end if
                pp.close: set pp = nothing
            end function        
        %>    
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->      

        <br />
        
        <div style="width: 95%; margin: auto;">
            <table style="width: 100%;">
                <tr>
                    <td style="width: 30%;">
                        <span style="font-size: 18px;"><%= "Asignar Roles a " & NombreUsuario(Usuario) %></span>
                    </td>

                    <td style="width: 70%; text-align: right;">
                        <button class='form-btn rojo normal'  type='button' onclick="volver()">Cancelar</button>&nbsp;&nbsp;
                        <button class='form-btn verde normal' type='button' onclick="grabar()">Actualizar</button>&nbsp;&nbsp;                                    
                    </td>
                </tr>
            </table>   
        </div>

        <div style="width: 98%; margin: auto;">
            <form id="formulario"  name="formulario" method="post" action="editar_usuario_grabar_roles.asp">
                <input id="codigo" name="codigo" type="text" value="<%= usuario %>" class="no-ver" />

                <div class="main main-scroll">
                    <div class="line label-top">
                        <label class="label small">Roles</label>
                        <div class="label full section">
                            <div class="tabla-wrapper">
                                <table class="tabla tabla-blue">
                                    <%
                                        sqlString = "SELECT r.rolCodigo, r.rolNombre, r.rolDescripcion," & _
                                                            " CASE WHEN r.TipoRol = 1 THEN 'Rol' ELSE 'Anti-Rol' END AS Tipo," & _
                                                            " CASE WHEN ru.CodigoUsuario IS NULL THEN 0 ELSE 1 END AS RolUsu " & _
                                                        "FROM dbo.seg_Roles AS r " & _
                                            "LEFT OUTER JOIN (SELECT CodigoRol, CodigoUsuario " & _
                                                                "FROM dbo.seg_RolesUsuarios " & _
                                                                "WHERE (CodigoUsuario = '" & Usuario & "')) AS ru " & _
                                                        "ON r.rolCodigo = ru.CodigoRol " & _
                                                    "ORDER BY r.rolNombre;"

                                        set tt = cc.execute(sqlString)
                                            if not (tt.bof or tt.eof) then
                                                do
                                                    lista = procRoles(tt("rolCodigo"))

                                                    %>
                                                        <tr>
                                                            <td style="padding: 10px;" title="<%= lista %>">
                                                                <input type="checkbox" id="<%= tt("rolCodigo") %>" name="<%= tt("rolCodigo") %>" value="1" <%
                                                                    if tt("RolUsu") = 1 then response.write " checked"
                                                                %>/>
                                                            </td>
                                                            <td style="padding: 10px; text-align: left !important;" title="<%= lista %>"><%= tt("rolNombre") %></td>
                                                            <td style="padding: 10px; text-align: left !important;" title="<%= lista %>"><%= tt("rolDescripcion") %></td>
                                                            <td style="padding: 10px; text-align: left !important;" title="<%= lista %>"><%= tt("Tipo") %></td>
                                                        </tr>
                                                    <%

                                                    tt.MoveNext
                                                loop until tt.eof
                                            end if
                                        tt.close: set tt = nothing
                                    %>
                                </table>    
                            </div>                                                    
                        </div>
                    </div>
                </div>
            </form>
        </div>

        <br /><br />

        <script type="text/javascript">
            function volver() {
                var vinculo = "lista.asp?o=<%= ordenadoPor %>";
                window.location.href = vinculo;
            }

            function grabar() {
                document.getElementById("formulario").submit();
            }
        </script>
    </body>

    <% cc.close: set cc = nothing %>  
    <!-- #include virtual = "/core/includes/kernel/close.inc" -->      
</html>