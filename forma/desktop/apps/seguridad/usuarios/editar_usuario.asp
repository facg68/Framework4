<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Editar Credenciales del Usuario</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->    

        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0020"
            SysLockOut

            '
            ' Init()
            '
            dim cc, tt, sqlString, usuario, ordenadoPor

            usuario = Request.querystring("u")
            ordenadoPor = Request.querystring("o")

            sqlString = "SELECT * " & _
                        "FROM seg_Usuarios " & _
                        "WHERE usuCodigo = '" & usuario & "';"

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")
            set tt = cc.execute(sqlString)

            '
            ' Funciones y Procedimientos
            '

            function FechaFormulario(FechaServer)
                dim d, m, a

                d = right("00" & day(FechaServer), 2)
                m = right("00" & month(FechaServer), 2)
                a = year(FechaServer)

                FechaFormulario = d & "/" & m & "/" & a
            end function

            sub PresentarSnippets()
                dim ss, ssSQl

                ssSQL = "SELECT s.codUsuario, s.codSistema, s.codProceso, s.snippet, s.snippetActivo, p.proNombre " & _
                        "FROM seg_Usuarios_Snippets AS s " & _
                    "INNER JOIN seg_Procesos AS p " & _
                            "ON s.codSistema = p.proSistema " & _ 
                        "AND s.codProceso = p.proCodigo " & _
                        "WHERE s.codUsuario = '" & usuario & "';"

                set ss = cc.execute(ssSQL)

                if not (ss.bof or ss.eof) then
                    Do
                        response.write "<p>"
                            response.write "<input type='checkbox' id='" & ss("Snippet") & "' name='" & ss("Snippet") & "' value='1'"
                                if ss("snippetActivo") = 1 then response.write " checked "
                            response.write ">"

                            response.write "<label>&nbsp;&nbsp;" & ss("proNombre") & "</label>"
                        response.write "</p>"                             

                        ss.MoveNext
                    Loop Until ss.eof
                end if

                ss.close: set ss = nothing
            end sub

            sub PresentarListaShortcuts()
                dim ss, ssSQl

                ssSQL = "exec dbo.seg_Shortcuts_Activos '" & usuario & "';"

                set ss = cc.execute(ssSQL)

                if not (ss.bof or ss.eof) then
                    Do
                        response.write "<p>"
                            response.write "<input type='checkbox' id='" & ss("Proceso") & "' name='" & ss("Proceso") & "' value='1'"
                                if ss("Activo") = 1 then response.write " checked "
                            response.write ">"

                            response.write "<label>&nbsp;&nbsp;" & ss("Nombre") & "</label>"
                        response.write "</p>"     

                        ss.MoveNext
                    Loop Until ss.eof
                end if

                ss.close: set ss = nothing
            end sub               

            function hPage()
                dim t, cmdString, con

                cmdString = "SELECT usuHomePage FROM seg_Usuarios WHERE usuCodigo = '" & usuario & "';"

                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")
                set t = con.execute(cmdString)

                if not (t.bof or t.eof) then
                    if len(trim(t("usuHomePage"))) > 0 then
                    hPage = "/apps/desktop/" & t("usuHomePage") & ".asp"
                    else
                    hPage = "/core/desktop.asp"
                    end if
                else
                    hPage = Application("DefPage")
                end if

                t.close: set t = nothing        
                con.close: set con = nothing
            end function   

            function NombreUsuario()     
                dim tu

                set tu = cc.execute("SELECT usuNombre FROM seg_Usuarios WHERE usuCodigo ='" & usuario & "';")
                    NombreUsuario = tu("usuNombre")
                tu.close: set tu = nothing
            end function

            function TieneSnippets()
                sqlString = "SELECT CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END AS Cuantos " & _
                                "FROM dbo.seg_Usuarios_Snippets AS s " & _
                                "WHERE (codUsuario = '" & Usuario & "');"   

                set ptt = cc.execute(sqlString)  
                    TieneSnippets = ptt("Cuantos") 
                ptt.close: set ptt = nothing                                   
            end function

            function TieneVariables()
                sqlString = "SELECT COUNT(*) AS Cuantas " & _
                            "FROM dbo.seg_Parametros AS p INNER JOIN dbo.seg_Usuarios_Parametros AS up " & _ 
                            "ON p.Parametro = up.Parametro AND p.Sistema = up.Sistema " & _
                            "WHERE (up.Usuario = '" & Usuario & "') AND (p.Exponer = 1);"    

                set ptt = cc.execute(sqlString)  
                    TieneVariables = ptt("Cuantas") 
                ptt.close: set ptt = nothing                                   
            end function

            function TieneShortcuts()
                sqlString = "SELECT CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END AS Cuantos " & _
                                "FROM dbo.seg_PermisosUsuarios AS pu " & _
                        "INNER JOIN dbo.seg_Procesos AS pr " & _
                                "ON pu.Sistema = pr.proSistema " & _
                                "AND pu.Proceso = pr.proCodigo " & _
                                "WHERE (pu.Usuario = '" & Usuario & "') " & _
                                "AND (pr.Shortcut = 1);"

                set ptt = cc.execute(sqlString)  
                    TieneShortcuts = ptt("Cuantos")
                ptt.close: set ptt = nothing     
            end function           
        %>    

        <style>
            p {
                margin-bottom: 10px;
                line-height: 1.5;
            }  
        </style>
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->      

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                 Editar Usuario [<%= usuario %>]
            </div>
            
            <div style="flex: 0 0 50%; text-align: right;">
                <button class='form-btn rojo normal'  type='button' onclick="volver()">Cancelar</button>&nbsp;&nbsp;
                <button class='form-btn naranja normal' type='button' onclick="password('<%= NombreUsuario() %>')">Password</button>&nbsp;&nbsp;
                <button class='form-btn verde normal' type='button' onclick="grabar()">Grabar</button>&nbsp;&nbsp;             
            </div>
        </div>                     


        <form id="formulario"  name="formulario" method="post" action="editar_usuario_grabar.asp">
            <input id="codigo" name="codigo" type="text" value="<%= usuario %>" class="no-ver"/>

            <div class="main main-scroll">
                <div class="line">
                    <label class="label large">Nombre</label>
                    <input class="field normal" id="usuNombre" name="usuNombre" type="text" value="<%= tt("usuNombre") %>" />
                </div>

                <div class="line">
                    <label class="label large">Cargo</label>
                    <input class="field large" id="usuCargo" name="usuCargo" type="text" value="<%= tt("usuCargo") %>" />
                </div>                  

                <div class="line">
                    <label class="label large">Correo</label>
                    <input class="field large" id="usuCorreo" name="usuCorreo" type="email" value="<%= tt("usuCorreo") %>" />
                </div>

                <div class="line">
                    <label class="label large">Fecha de Nacimiento</label>
                    <input class="field tiny" id="usuFechaNacimiento" name="usuFechaNacimiento" type="text" value="<%= FechaFormulario(tt("usuFechaNacimiento")) %>" placeholder="dd/mm/aaaa" />
                </div>

                <div class="line">
                    <label class="label large">Teléfono o Exensión</label>
                    <input class="field small" id="usuTelefono" name="usuTelefono" type="email" value="<%= tt("usuTelefono") %>" />
                </div>

                <div class="line">
                    <label class="label large">Estado del Usuario</label>

                    <%
                        response.write "<select class='field normal' name='usuEstado' id='usuEstado'>"
                            response.write "<option value='1' "
                                if tt("usuEstado") = "1" then response.write " selected"
                            response.write ">Activo</option>"

                            response.write "<option value='0' "
                                if tt("usuEstado") = "0" then response.write " selected"
                            response.write ">Desactivado</option>"
                        response.write "/<select>"
                    %>
                </div>                                            

                <div class="line">
                    <label class="label large">Acción del Botón &#9733;</label>

                    <%
                        dim cboConn, cboSqlString, cboTable

                        cboSqlString = "SELECT Proceso, Pagina " & _
                                        "FROM (SELECT s.sysNombre + ': ' + p.proNombre AS Proceso, s.sysCodigo + '/' + u.proAction AS Pagina " & _
                                                "FROM dbo.seg_Procesos AS p " & _
                                                "INNER JOIN dbo.seg_PermisosUsuarios AS u " & _
                                                "ON p.proSistema = u.Sistema " & _
                                                "AND p.proCodigo = u.Proceso " & _
                                                "INNER JOIN dbo.seg_Sistemas AS s " & _
                                                "ON p.proSistema = s.sysCodigo " & _
                                        "WHERE (p.proHomePage = 1) " & _
                                        "AND (u.Usuario = '" & usuario  & "')) AS t " & _
                                        "ORDER BY Proceso;"

                        set cboConn = Server.CreateObject("ADODB.Connection")
                        cboConn.open Application("Conn")
                        set cboTable = cboConn.execute(cboSqlString)

                        if not(cboTable.bof or cboTable.eof) then
                            response.write "<select class='field xxl' name='usuHomePage' id='usuHomePage'>"
                                response.write "<option value=''"
                                    if tt("usuHomePage") = "" then 
                                        response.write " selected"
                                    end if
                                response.write ">&nbsp;[&nbsp;Modo Escritorio&nbsp;]&nbsp;</option>"

                                do
                                    response.write "<option value='" & cboTable("Pagina") & "' " 
                                        if tt("usuHomePage") = cboTable("Pagina") then 
                                            response.write " selected"
                                        end if
                                    response.write ">" & cboTable("Proceso")  & "</option>"
                                
                                    cboTable.MoveNext
                                loop until cboTable.eof

                            response.write "/<select>"
                        end if

                        cboTable.close: set cboTable = nothing
                        cboConn.close: set cboConn = nothing
                    %>
                </div>

                <div class="line">
                    <label class="label large">Parametros</label>
                    <label class="label full section" style="text-align: left !important;">
                        <p>
                            <input type="checkbox" id="usuVerSaldos" name="usuVerSaldos" value="1" <% if tt("usuVerSaldos") = 1 then response.write "checked" %>/>
                            <label for="usuVerSaldos">&nbsp;&nbsp;Ver Saldos de Presupuesto en Encabezado</label>
                        </p>

                        <p>
                            <input type="checkbox" id="usuIniciarSinEncabezado" name="usuIniciarSinEncabezado" value="1" <% if tt("usuIniciarSinEncabezado") = 1 then response.write "checked" %>/>
                            <label for="usuIniciarSinEncabezado">&nbsp;&nbsp;Iniciar las Opciones con el Menú Oculto</label>
                        </p>

                        <p>
                            <input type="checkbox" id="usuCargarSnippets" name="usuCargarSnippets" value="1" <% if tt("usuCargarSnippets") = 1 then response.write "checked" %>/>
                            <label for="usuCargarSnippets">&nbsp;&nbsp;Cargar Snippets en el Modo Escritorio</label>
                        </p>

                        <p>
                            <input type="checkbox" id="usuRandomWallpaper" name="usuRandomWallpaper" value="1" <% if tt("usuRandomWallpaper") = 1 then response.write "checked" %>/>
                            <label for="usuCargarSnippets">&nbsp;&nbsp;Fondo Aleatorio (Modo Escritorio)</label>
                        </p>
                    </label>
                </div>

                <% if TieneShortcuts() = 1 then %>
                    <div class="line">
                        <label class="label large">Vinculos</label>
                        <label class="label full section" style="text-align: left !important;">
                            <%
                                PresentarListaShortcuts()
                            %>                    
                        </label>
                    </div>
                <% end if %>

                <% if TieneSnippets() = 1 then %>                                                
                    <div class="line">
                        <label class="label large">Snippets</label>
                        <label class="label full section" style="text-align: left !important;">
                            <%
                                PresentarSnippets()
                            %>                    
                        </label>
                    </div>

                    <div class="line">
                        <label class="label large">Transparencia (Snippets)</label>
                        <label class="label full section">
                            <input style="width: 90%;" type="range" min="60" max="100" 
                                    value="<%= tt("snippetsOpacidad") %>" 
                                    id="snippetsOpacidad" name="snippetsOpacidad">
                            
                            <label style="width: 10%;" id="OpacidadValor"></label>  

                            <script>
                                var slider = document.getElementById("snippetsOpacidad");
                                var output = document.getElementById("OpacidadValor");

                                output.innerHTML = slider.value;

                                slider.oninput = function() {
                                    output.innerHTML = this.value;
                                }
                            </script>                                        
                        </label>
                    </div>
                <% end if %>    
            </div>            
        </form>

        <br />

        <script type="text/javascript">
            function volver() {
                var vinculo = "lista.asp?o=<%= ordenadoPor %>";
                window.location.href = vinculo;
            }

            function password(nombre) {
                var confirmacion = confirm("Desea resetear el password del usuario " + nombre + "?");
                var vinculo = "editar_usuario_password.asp?u=<%= usuario %>&o=<%= ordenadoPor %>";          

                if (confirmacion) {     
                    window.location.href = vinculo;
                };              
            }

            function grabar() {
                document.getElementById("formulario").submit();
            }

            mask(document.getElementById('usuFechaNacimiento'), ['99/99/9999']);
            mask(document.getElementById('snippetsOpacidad'),   ['999']);
        </script>
    </body>

    <%
        tt.close: set tt = nothing
        cc.close: set cc = nothing
    %>    
    <!-- #include virtual = "/core/includes/kernel/close.inc" -->
</html>