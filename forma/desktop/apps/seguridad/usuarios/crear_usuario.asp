<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Crear Nuevo Usuario</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->        
        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0020"
            SysLockOut

            '
            ' Init()
            '
            dim cc, tt, pp, sqlString, usuario, ordenadoPor, linea, lista

            set cc = Server.CreateObject("ADODB.Connection")
            cc.Open Application("Conn")

            ordenadoPor = Request.querystring("o")

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

    <body reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->  

        <br />    

        <div style="width: 95%; margin: auto;">
            <table style="width: 100%;">
                <tr>
                    <td style="width: 30%; font-size: 24px;">
                        Crear Usuario
                    </td>

                    <td style="width: 70%; text-align: right;">
                        <button class='form-btn rojo normal'  type='button' onclick="volver()">Cancelar</button>&nbsp;&nbsp;
                        <button class='form-btn verde normal' type='button' onclick="grabar()">Grabar</button>&nbsp;&nbsp;
                    </td>
                </tr>
            </table>  
        </div>

        <div style="width: 95%; margin: auto;">
            <form id="formulario"  name="formulario" method="post" action="grabar_usuario_nuevo.asp">
                <div class="main main-scroll">
                    <div class="line">
                        <label class="label normal">Codigo</label>
                        <input class="field small" id="usuCodigo" name="usuCodigo" type="text" required />
                    </div>

                    <div class="line">
                        <label class="label normal">Nombre</label>
                        <input class="field xl" id="usuNombre" name="usuNombre" type="text" required />
                    </div>

                    <!--
                        Lista de Roles
                    -->

                        <div class="line label-top">
                            <label class="label normal">Roles</label>
                            <div class="label full section">
                                <div class="tabla-wrapper">
                                    <table class="tabla tabla-green">
                                        <%
                                            sqlString = "SELECT rolCodigo, rolNombre, rolDescripcion," & _
                                                                " CASE WHEN TipoRol = 1 THEN 'Rol' ELSE 'Anti-Rol' END AS Tipo " & _
                                                            "FROM seg_Roles " & _
                                                        "ORDER BY rolNombre;"

                                            set tt = cc.execute(sqlString)
                                                if not (tt.bof or tt.eof) then
                                                    do
                                                        lista = procRoles(tt("rolCodigo"))

                                                        %>
                                                            <tr>
                                                                <td title="<%= lista %>"><input type="checkbox" id="<%= tt("rolCodigo") %>" name="<%= tt("rolCodigo") %>" value="1" /></td>
                                                                <td title="<%= lista %>"><%= tt("rolNombre") %></td>
                                                                <td title="<%= lista %>"><%= tt("rolDescripcion") %></td>
                                                                <td title="<%= lista %>"><%= tt("Tipo") %></td>
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

                    <!--
                        Fin de la Lista de Roles
                    -->
                </div>
            </form>
        </div>

        <br />

        <script type="text/javascript">
            function volver() {
                var vinculo = "lista.asp?o=<%= ordenadoPor %>";
                window.location.href = vinculo;
            }

            function grabar() {
                document.getElementById("formulario").submit();
            }
        </script>

        <% cc.close: set cc = nothing %>
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>