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
            dim cc, tt, sqlString, usuario

            usuario = Request.Cookies("usuario")

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
                        hPage = "/forma/desktop/apps/" & t("usuHomePage") & ".asp"
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

            Function ConsultaSQL(Sistema, Parametro) 
                dim sqlString

                sqlString = "SELECT ConsultaSQL AS Consulta FROM seg_Parametros WHERE (Sistema = '" & Sistema & "' AND Parametro = '" & Parametro & "');"

                set ptt = cc.execute(sqlString)   
                    ConsultaSQL = ptt("Consulta") 
                    ConsultaSQL = Replace(ConsultaSQL, "|", "'")
                ptt.close: set ptt = nothing                
            End Function   

            sub VariablesAsignadas()   
                dim conteo   
                %>
                    <table style="width: 100%;" style="padding: 0px;"> 
                        <%
                            sqlString = "SELECT up.Sistema, up.Parametro, p.TipoParametro, s.sysNombre, p.Descripcion, up.Valor, up.Sistema + '__' + up.Parametro AS Variable " & _
                                            "FROM dbo.seg_Parametros AS p " & _
                                    "INNER JOIN dbo.seg_Sistemas AS s " & _
                                            "ON p.Sistema = s.sysCodigo " & _ 
                                    "INNER JOIN dbo.seg_Usuarios_Parametros AS up " & _ 
                                            "ON p.Parametro = up.Parametro " & _ 
                                            "AND p.Sistema = up.Sistema " & _
                                            "WHERE (up.Usuario = '" & Usuario & "') " & _
                                            "AND (p.Exponer = 1) " & _
                                        "ORDER BY p.Descripcion;"

                            set cbox = cc.execute(sqlString)
                                if not (cbox.bof or cbox.eof) then
                                    conteo = 0

                                    Do       
                                        conteo = conteo + 1

                                        response.write "<tr style='border: none;'>"
                                            response.write "<td style='text-align: left; padding: 10px; width: 70%;'>"
                                                response.write RIGHT("00" & conteo, 2) & ". " & cbox("Descripcion")
                                            response.write "</td>"

                                            response.write "<td style='text-align: left; padding: 10px; width: 30%;'>"
                                                Select Case cbox("TipoParametro")
                                                    Case "1"
                                                        '
                                                        ' Es un Permiso... El Campo tiene un valor de 1 y está escondido
                                                        '
                                                        response.write "<input class='field full' id='" & cbox("Variable") & " name=" & cbox("Variable") & " type='text' "
                                                        response.write "value=" & cbox("Valor") & " style='visibility: hidden;' />"
                                                    Case "2"
                                                        '
                                                        ' El Campo es una Variable... Presentamos el input-box
                                                        '
                                                        %>
                                                            <input class="field full" id="<%= cbox("Variable") %>" name="<%= cbox("Variable") %>" type="text" value="<%= cbox("Valor") %>" />
                                                        <%
                                                    Case "3"
                                                        '
                                                        ' La Variable es un Campo "Si/No". Presentamos un Check-Box
                                                        '
                                                        %>
                                                            <input type="radio" id="<%= cbox("Variable") & "1" %>" name="<%= cbox("Variable") %>" value="1" <%
                                                                if cbox("Valor") = "1" then response.write " checked " %>>
                                                            &nbsp;&nbsp;
                                                            <label style="font-family: Verdana; font-size: 14px;" for="Si">Si</label>

                                                            &nbsp;&nbsp;&nbsp;&nbsp;

                                                            <input type="radio" id="<%= cbox("Variable") & "0" %>" name="<%= cbox("Variable") %>" value="0" <%
                                                                if cbox("Valor") = "0" then response.write " checked " %> >
                                                            &nbsp;&nbsp;
                                                            <label style="font-family: Verdana; font-size: 14px;" for="No">No</label>
                                                        <%

                                                    Case "4"
                                                        '
                                                        ' Es una lista... Presentamos un Combo-Box
                                                        '
                                                        set lista = cc.execute("SELECT * FROM seg_Parametros_Valores WHERE Sistema = '" & cbox("Sistema") & "' AND Parametro = '" & cbox("Parametro") & "' ORDER BY Descripcion;")

                                                        if not (lista.bof or lista.eof) then
                                                            %><select name="<%= cbox("Variable") %>" id="<%= cbox("Variable") %>" class="field full"><%
                                                                do
                                                                    response.write "<option value='" & lista("Valor") & "'"
                                                                        if cbox("Valor") = lista("Valor") then
                                                                            response.write " selected"
                                                                        end if
                                                                    response.write ">" & lista("Descripcion") & "</option>"
                                                                    
                                                                    lista.MoveNext
                                                                loop until (lista.eof)
                                                            response.write "</select>"
                                                        end if

                                                        lista.close: set lista = nothing

                                                    Case "5"
                                                        '
                                                        ' Es una Consulta... Presentamos un Combo-Box
                                                        '
                                                        Consulta = ConsultaSQL(cbox("Sistema"), cbox("Parametro"))

                                                        set lista = cc.execute(Consulta)

                                                        if not (lista.bof or lista.eof) then
                                                            %><select name="<%= cbox("Variable") %>" id="<%= cbox("Variable") %>" class="field full"><%
                                                                do
                                                                    response.write "<option value='" & lista("Codigo") & "'"
                                                                        if cbox("Valor") = lista("Codigo") then
                                                                            response.write " selected"
                                                                        end if
                                                                    response.write ">" & lista("Valor") & "</option>"
                                                                    
                                                                    lista.MoveNext
                                                                loop until (lista.eof)
                                                            response.write "</select>"
                                                        end if

                                                        lista.close: set lista = nothing 

                                                    Case "6"
                                                        '
                                                        ' El Campo es un selector de color... Presentamos la rueda de colores
                                                        '
                                                        %>
                                                            <input class="field full" id="<%= cbox("Variable") %>" name="<%= cbox("Variable") %>" type="color" value="<%= cbox("Valor") %>" />
                                                        <%

                                                    Case 7
                                                        %>
                                                            <label style="width: 100%;">
                                                                <input class="field full" style="width:100%" type="range" min="0" max="100" 
                                                                    value="<%= cbox("Valor") %>" 
                                                                    id="<%= cbox("Variable") %>" name="<%= cbox("Variable") %>">
                                                            </label>                                                      
                                                        <%
                                                End Select
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

        <style>
            p {
                margin-bottom: 10px;
                line-height: 1.5;
            }          
        </style>
    </head>

    <body plantilla="normal" reserva="140">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->      

        <br />

        <div style="width: 95%; margin: auto;">
            <table style="width: 100%;">
                <tr>
                    <td style="width: 30%; font-size: 24px;">
                        Editar Usuario [<%= usuario %>]
                    </td>

                    <td style="width: 70%; text-align: right;">
                        <button class='form-btn rojo normal'  type='button' onclick="volver()">Cancelar</button>&nbsp;&nbsp;
                        <button class='form-btn naranja normal' type='button' onclick="password('<%= NombreUsuario() %>')">Password</button>&nbsp;&nbsp;
                        <button class='form-btn verde normal' type='button' onclick="grabar()">Grabar</button>&nbsp;&nbsp;
                    </td>
                </tr>
            </table>            


            <form id="formulario"  name="formulario" method="post" action="/core/includes/profile_grabar.asp">
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

                    <div class="line label-top">
                        <label class="label large">Parametros</label>
                        <label class="label full section" style="text-align: left !important;">
                            <p>
                                <input type="checkbox" id="usuVerSaldos" name="usuVerSaldos" value="1" <% if tt("usuVerSaldos") = 1 then response.write "checked" %>/>
                                <label>&nbsp;&nbsp;Ver Saldos de Presupuesto en Encabezado</label>
                            </p>

                            <p>
                                <input type="checkbox" id="usuIniciarSinEncabezado" name="usuIniciarSinEncabezado" value="1" <% if tt("usuIniciarSinEncabezado") = 1 then response.write "checked" %>/>
                                <label>&nbsp;&nbsp;Iniciar las Opciones con el Menú Oculto</label>
                            </p>

                            <p>
                                <input type="checkbox" id="usuCargarSnippets" name="usuCargarSnippets" value="1" <% if tt("usuCargarSnippets") = 1 then response.write "checked" %>/>
                                <label>&nbsp;&nbsp;Cargar Snippets en el Modo Escritorio</label>
                            </p>

                            <p>
                                <input type="checkbox" id="usuRandomWallpaper" name="usuRandomWallpaper" value="1" <% if tt("usuRandomWallpaper") = 1 then response.write "checked" %>/>
                                <label>&nbsp;&nbsp;Fondo Aleatorio (Modo Escritorio)</label>
                            </p>
                        </label>
                    </div>

                    <% if TieneShortcuts() = 1 then %>
                        <div class="line label-top">
                            <label class="label large">Vinculos</label>
                            <label class="label full section" style="text-align: left !important;">
                                <%
                                    PresentarListaShortcuts()
                                %>                    
                            </label>
                        </div>
                    <% end if %>

                    <% if TieneSnippets() = 1 then %>                                                
                        <div class="line label-top">
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

                    <% if TieneVariables() > 0 then %>
                        <div class="line label-top">
                            <label class="label large">Variables</label>
                            <label class="label full section" style="text-align: left !important;">
                                <%
                                    VariablesAsignadas()
                                %>                    
                            </label>
                        </div>
                    <% end if %>
                </div>            
            </form>
        </div>

        <br /><br />

        <script type="text/javascript">
            function volver() {
                var vinculo = "<%= hPage() %>";
                window.location.href = vinculo;
            }

            function password() {
                var vinculo = "/core/includes/profile_password.asp";
                window.location.href = vinculo;
            }

            function grabar() {
                document.getElementById("formulario").submit();
            }

            mask(document.getElementById('usuFechaNacimiento'), ['99/99/9999']);
        </script>
    </body>

    <%
        tt.close: set tt = nothing
        cc.close: set cc = nothing
    %>    
    <!-- #include virtual = "/core/includes/kernel/close.inc" -->
</html>