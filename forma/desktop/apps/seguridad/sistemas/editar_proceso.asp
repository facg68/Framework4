<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Editar Procesos</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0070"
            SysLockOut

            dim pCon, ptt, p, sqlString, Sistema, Proceso, ordenadoPor
            dim Nombre, Activo, MenuItem, MenuIndice, Icon, Root, Action, proActionParam, pHomePage, Snippet, Shortcut, Externo

            Sistema = request.querystring("s")
            Proceso = request.querystring("p")
            ordenadoPor = request.querystring("op")

            set pCon = Server.CreateObject("ADODB.Connection")
            pCon.open Application("Conn")

            function NombreProceso(Sistema, Proceso)
                dim sqlString

                sqlString = "SELECT s.sysNombre AS Sistema, p.proNombre AS Proceso " & _
                                "FROM seg_Procesos AS p " & _
                        "INNER JOIN seg_Sistemas AS s " & _
                                "ON p.proSistema = s.sysCodigo " & _
                                "WHERE (p.proSistema = '" & Sistema & "') " & _
                                "AND (p.proCodigo = '" & Proceso & "');"

                set ptt = pCon.execute(sqlString)   
                    NombreProceso = ptt("Sistema") & "&nbsp;&nbsp;(&nbsp;" & ptt("Proceso") & "&nbsp;)"
                ptt.close: set ptt = nothing
            end function    

            function NombreSistema(Sistema)
                dim sqlString

                sqlString = "SELECT sysNombre FROM seg_Sistemas WHERE (sysCodigo ='" & Sistema & "');"

                set ptt = pCon.execute(sqlString)   
                    NombreSistema = ptt("sysNombre") 
                ptt.close: set ptt = nothing
            end function          
        %>
        <style>
            body { overflow: none; }
        </style>
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->   

        <%
            if Proceso <> "*" then
                set ptt = pCon.execute("SELECT * FROM seg_Procesos WHERE (proSistema = '" & Sistema & "') AND (proCodigo = '" & Proceso & "');")
                    Nombre = ptt("proNombre")
                    Activo = ptt("proActivo")
                    MenuItem = ptt("proMenuItem")
                    MenuIndice = ptt("proMenuIndice")
                    Icon = ptt("proIcon")
                    Root = ptt("proRoot")
                    Action = ptt("proAction")
                    proActionParam= ptt("proActionParam")
                    pHomePage = ptt("proHomePage")
                    Snippet = ptt("Snippet")
                    Shortcut = ptt("Shortcut")
                    Movil = ptt("Movil")
                    Externo = ptt("Externo")
                ptt.close: set ptt = nothing
            end if
        %>

        <div style="width: 95%; margin: auto;">
            <br />

            <table style="width: 100%;">
                <tr>
                    <td style="width: 30%; font-size: 24px;">
                        <%
                            if Proceso = "*" then
                                response.write "Nuevo Proceso  [" & NombreSistema(Sistema) & "]"
                            else
                                response.write NombreProceso(Sistema, Proceso)
                            end if
                        %>
                    </td>

                    <td style="width: 70%; text-align: right;">
                        <button class='form-btn rojo normal'  type='button' onclick="volver()">Cancelar</button>&nbsp;&nbsp;
                        <button class='form-btn verde normal' type='button' onclick="grabar()">Grabar</button>&nbsp;&nbsp;
                    </td>
                </tr>
            </table>    

            <form id="formulario"  name="formulario" method="post" action="grabar_proceso.asp">
                <input id="Sistema" name="Sistema" type="text" value="<%= Sistema %>" class="no-ver" />
                <input id="ordenadoPor" name="ordenadoPor" type="text" value="<%= ordenadoPor %>" class="no-ver" />

                <div class="main main-scroll">
                    <div class="line">
                        <% if Proceso = "*" then %>
                            <input id='nuevo' name='nuevo' type='text' value='1' class="no-ver" />
                            
                            <label class="label large" for="Codigo">Codigo</label>                               
                            <input class="field small"  id="Codigo" name="Codigo" type="text" required />
                        <% else %>
                            <input id='nuevo' name='nuevo' type='text' value='0' class="no-ver" />
                            <input id="Codigo" name="Codigo" type="text" value="<%= Proceso %>" class="no-ver" />

                            <label class="label large" for="Codigo">Codigo</label>
                            <input class="field small" id="dispCodigo" name="dispCodigo" type="text" value='<%= Proceso %>' disabled />
                        <% end if %>
                    </div>

                    <div class="line">
                        <label class="label large">Nombre</label>
                        <input class= "field xxl" id="Nombre" name="Nombre" type="text" required <% if Proceso <> "*" then response.write "value='" & Nombre & "'" %> />
                    </div>

                    <div class="line">
                        <label class="label large">Página Proceso (Web)</label>
                        <input class="field xxl" id="Action" name="Action" type="text" required <% if Proceso <> "*" then response.write "value='" & Action & "'" %> />
                    </div>     

                    <div class="line">
                        <label class="label large">Apéndice (Parámetros)</label>
                        <input class="field large" id="proActionParam" name="proActionParam" type="text" required <% if Proceso <> "*" then response.write "value='" & proActionParam & "'" %> />
                    </div>                                         

                    <div class="line">
                        <label class="label large">
                            <img src="imagenes/menu_vacio.png" style="border: none; width: 25px; height: 25px;">&nbsp;Opción de Menú
                        </label>

                        <select class="field normal" name="MenuItem" id="MenuItem">
                            <option value="0" <% if Proceso <> "*" then 
                                                    if MenuItem = 0 then response.write "selected"
                                                end if
                                            %> >Encabezado</option>

                            <option value="1" <% if Proceso <> "*" then 
                                                    if MenuItem = 1 then response.write "selected"
                                                else
                                                    response.write "selected"
                                                end if
                                            %> >Opcion de Menu</option>
                        </select>
                    </div>

                    <div class="line">
                        <label class="label large">Ordenamiento en Menú</label>
                        <input class="field small" id="MenuIndice" name="MenuIndice" type="text" required <% if Proceso <> "*" then response.write "value='" & MenuIndice & "'" %> />
                    </div>

                    <div class="line">
                        <label class="label large">Pertenece A</label>
                        <select class="field large" name="Root" id="Root">
                            <option value = NULL>&nbsp;</option>
                            <%
                                set ptt = pCon.execute("SELECT proCodigo AS Codigo, proNombre AS Nombre " & _
                                                        "FROM seg_Procesos " & _
                                                        "WHERE (proAction IS NULL) " & _
                                                            "AND (proSistema = '" & Sistema & "') " & _
                                                    "ORDER BY proNombre;")

                                    if not (ptt.bof or ptt.eof) then
                                        do
                                        response.write "<option value='" & ptt("Codigo") & "'"

                                        if Proceso <> "*" then 
                                            if Root = ptt("Codigo") then 
                                            response.write "selected"
                                            end if
                                        end if

                                        response.write ">" & ptt("Nombre") & "</option>"

                                        ptt.MoveNext
                                        loop until ptt.eof
                                    end if  

                                ptt.close: set ptt = nothing                                  
                            %>
                        </select>
                    </div>   

                    <div class="line">
                        <label class="label large">
                            <img src="imagenes/home_page.png" style="border: none; width: 25px; height: 25px;">&nbsp;HomePage
                        </label>

                        <select class="field normal" name="pHomePage" id="pHomePage">
                            <option value="0" <% if Proceso <> "*" then 
                                                    if pHomePage = 0 then response.write "selected"
                                                else
                                                    response.write "selected"
                                                end if
                                            %> >&nbsp;</option>

                            <option value="1" <% if Proceso <> "*" then 
                                                    if pHomePage = 1 then response.write "selected"
                                                end if
                                            %> >Es un HomePage</option>
                        </select>
                    </div>                            

                    <div class="line">
                        <label class="label large">
                            <img src="imagenes/snippet.png" style="border: none; width: 25px; height: 25px;">&nbsp;Snippet
                        </label>
                        
                        <input class="field normal"id="Snippet" name="Snippet" type="text" required <% if Proceso <> "*" then response.write "value='" & Snippet & "'" %> />
                    </div>                                      

                    <div class="line">
                        <label class="label large">
                            <img src="imagenes/vinculo.png" style="border: none; width: 25px; height: 25px;">&nbsp;Vinculo
                        </label>

                        <select class="field normal" name="Shortcut" id="Shortcut">
                            <option value="0" <% if Proceso <> "*" then 
                                                    if Shortcut = 0 then response.write "selected"
                                                else
                                                    response.write "selected"
                                                end if
                                            %> >&nbsp;</option>

                            <option value="1" <% if Proceso <> "*" then 
                                                    if Shortcut = 1 then response.write "selected"
                                                end if
                                            %> >Puede ser Anclado</option>
                        </select>
                    </div>  

                    <div class="line">
                        <label class="label large">
                            <img src="imagenes/movil.png" style="border: none; width: 25px; height: 25px;">&nbsp;Applet Movil
                        </label>

                        <select class="field normal" name="Movil" id="Movil">
                            <option value="0" <% if Proceso <> "*" then 
                                                    if Movil = 0 then response.write "selected"
                                                else
                                                    response.write "selected"
                                                end if
                                            %> >&nbsp;</option>

                            <option value="1" <% if Proceso <> "*" then 
                                                    if Movil = 1 then response.write "selected"
                                                end if
                                            %> >Tiene applet para moviles</option>
                        </select>
                    </div>                                            

                    <div class="line">
                        <label class="label large">Imagen de Vinculo</label>
                        <input class="field normal" id="Icon" name="Icon" type="text" required <% if Proceso <> "*" then response.write "value='" & Icon & "'" %> />
                    </div>

                    <div class="line">
                        <label class="label large">Obligatoriamente Externo</label>

                        <select class="field large" name="Externo" id="Externo">
                            <option value="0" <% if Proceso <> "*" then 
                                                    if Externo = 0 then response.write "selected"
                                                else
                                                    response.write "selected"
                                                end if
                                            %> >Corre dentro de la Extranet</option>

                            <option value="1" <% if Proceso <> "*" then 
                                                    if Externo = 1 then response.write "selected"
                                                end if
                                            %> >Necesita correr fuera de la Extranet</option>
                        </select>
                    </div>                         

                    <div class="line">
                        <label class="label large">Estado</label>
                        <select class="field normal" name="Activo" id="Activo">
                            <option value="0" <% if Proceso <> "*" then 
                                                    if Activo = 0 then response.write "selected"
                                                end if
                                            %> >Desactivado</option>

                            <option value="1" <% if Proceso <> "*" then 
                                                    if Activo = 1 then response.write "selected"
                                                else
                                                    response.write "selected"
                                                end if
                                            %> >Activo</option>
                        </select>
                    </div>
                </div>
            </form>
        </div>

        <br /><br />

        <script type="text/javascript">
            function volver() {
                var vinculo = "procesos.asp?s=<%= Sistema %>&o=<%= ordenadoPor %>";
                window.location.href = vinculo;
            }

            function grabar() {
                document.getElementById("formulario").submit();
            }          
        </script> 

        <% pCon.close: set pCon = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->    
    </body>
</html>