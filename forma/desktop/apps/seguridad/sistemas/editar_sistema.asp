<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">
        <title>Editar Sistema</title>

        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  

        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0060"
            SysLockOut

            dim con, t, p, sqlString, Sistema, ordenadoPor
            dim nombre, Descripcion, ClaseApp, IndiceOrdenamiento, Icono, sBitacora

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")            

            function TieneVariables(Sistema)
                dim ptt

                sqlString = "SELECT s.sysCodigo, ISNULL(P.Cuantos, 0) AS Variables " & _
                              "FROM dbo.seg_Sistemas AS s " & _
                   "LEFT OUTER JOIN (SELECT Sistema, COUNT(Parametro) AS Cuantos " & _
                                      "FROM dbo.seg_Parametros " & _
                                  "GROUP BY Sistema) AS P " & _
                                "ON s.sysCodigo = P.Sistema " & _
                             "WHERE (s.sysCodigo = '" & Sistema & "');"                                    

                set ptt = con.execute(sqlString)  
                    if not (ptt.bof or ptt.eof) then
                      TieneVariables = ptt("Variables") 
                    else
                      TieneVariables = 0
                    end if
                ptt.close: set ptt = nothing                                   
            end function     

            Function ConsultaSQL(Sistema, Parametro) 
                dim sqlString, ptt

                sqlString = "SELECT ConsultaSQL AS Consulta FROM seg_Parametros WHERE (Sistema = '" & Sistema & "' AND Parametro = '" & Parametro & "');"

                set ptt = con.execute(sqlString)   
                    ConsultaSQL = ptt("Consulta") 
                    ConsultaSQL = Replace(ConsultaSQL, "|", "'")
                ptt.close: set ptt = nothing   
            End Function     

            Function ConsultaSQLDefault(Sistema, Parametro) 
                dim sqlString, ptt

                sqlString = "SELECT ValorDefault FROM seg_Parametros WHERE (Sistema = '" & Sistema & "' AND Parametro = '" & Parametro & "');"

                set ptt = con.execute(sqlString)   
                    ConsultaSQLDefault = ptt("ValorDefault") 
                ptt.close: set ptt = nothing   
            End Function                      

            sub VariablesAsignadas(Sistema)   
                dim conteo  

                sqlString = "SELECT Sistema, Parametro, TipoParametro, Descripcion, ValorDefault AS Valor, Sistema + '__' + Parametro AS Variable " & _
                                "FROM dbo.seg_Parametros AS p " & _
                                "WHERE (Sistema = '" & Sistema & "') " & _
                            "ORDER BY Descripcion;"                 
                %>
                    <table style="width: 100%;" style="padding: 0px;"> 
                        <%
                            set cbox = con.execute(sqlString)
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
                                                        response.write "<input class='no-ver' id='" & cbox("Variable") & " name=" & cbox("Variable") & " type='text' "
                                                        response.write "value=" & cbox("Valor") & " />"
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
                                                        set lista = con.execute("SELECT * FROM seg_Parametros_Valores WHERE Sistema = '" & cbox("Sistema") & "' AND Parametro = '" & cbox("Parametro") & "' ORDER BY Descripcion;")

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

                                                        set lista = con.execute(Consulta)

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
                                                            <input class="field full" 
                                                                id="<%= cbox("Variable") %>" 
                                                                name="<%= cbox("Variable") %>" 
                                                                type="color" value="<%= cbox("Valor") %>" 
                                                                style="background-color: <%= cbox("Valor") %>; border: solid 1px rgb(200, 200, 200);"
                                                                onChange="reColor('<%= cbox("Variable") %>');"
                                                            />                                                            
                                                        <%

                                                    Case 7
                                                        %>
                                                            <input class="field full" type="range" min="0" max="100" 
                                                                   value="<%= cbox("Valor") %>" 
                                                                   id="<%= cbox("Variable") %>" name="<%= cbox("Variable") %>">                                                                                                               
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
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->   

        <%
            Sistema = request.querystring("s")
            ordenadoPor = request.querystring("op")

            if Sistema <> "*" then
                set t = con.execute("SELECT * FROM seg_Sistemas WHERE sysCodigo = '" & Sistema & "';")
                    Nombre = t("sysNombre")
                    Descripcion = t("sysDescripcion")
                    ClaseApp = t("sysWeb")
                    IndiceOrdenamiento = t("sysMenuIndice")
                    Icono = t("sysIcon")
                    sBitacora = t("sysBitacora")
                t.close: set t = nothing
            end if
        %>

        <div style="width: 95%; margin: auto;">
            <br />

            <table style="width: 100%;">
                <tr>
                    <td style="width: 30%; font-size: 24px;">
                        <%
                            if Sistema = "*" then
                                response.write "Nuevo Sistema"
                            else
                                response.write "Editar Sistema"
                            end if
                        %>
                    </td>

                    <td style="width: 70%; text-align: right;">
                        <button class='form-btn rojo normal'  type='button' onclick="volver()">Cancelar</button>&nbsp;&nbsp;
                        <button class='form-btn verde normal' type='button' onclick="grabar()">Grabar</button>&nbsp;&nbsp;
                    </td>
                </tr>
            </table>
        </div>

        <div style="width: 98%; margin: auto;">     
            <form id="formulario" name="formulario" method="post" action="grabar_sistema.asp">
                <input id="ordenadoPor" name="ordenadoPor" type="text" value="<%= ordenadoPor %>" class="no-ver" />     
                

                <div class="main main-scroll">
                    <div class="line">
                        <% if Sistema = "*" then %>
                            <label class="label normal">Codigo</label>
                            <input class="field small" id="Codigo" name="Codigo" type="text" required />

                            <input id='nuevo' name='nuevo' type='text' value='1' class="no-ver" />
                        <% else %>
                            <label class="label normal">Codigo</label>
                            <input class="field small" id="dispCodigo" name="dispCodigo" type="text" value='<%= Sistema %>' disabled />

                            <input id='nuevo' name='nuevo' type='text' value='0' class="no-ver" />
                            <input id="Codigo" name="Codigo" type="text" value="<%= Sistema %>" class="no-ver" />
                        <% end if %>
                    </div>

                    <div class="line">
                        <label class="label normal">Nombre</label>                       
                        <input class="field large"  id="Nombre" name="Nombre" type="text" required <% if Sistema <> "*" then response.write "value='" & Nombre & "'" %> />
                    </div>

                    <div class="line">
                        <label class="label normal">Descripcion</label>                        
                        <input class="field xl"  id="Descripcion" name="Descripcion" type="text" required <% if Sistema <> "*" then response.write "value='" & Descripcion & "'" %> />
                    </div>

                    <div class="line">
                        <label class="label normal">Tipo de Sistema</label>

                        <select class="field normal"  name="ClaseApp" id="ClaseApp" >
                            <option value="0" <% if Sistema <> "*" then 
                                                    if ClaseApp = 0 then response.write "selected"
                                                end if
                                            %> >Escritorio</option>

                            <option value="1" <% if Sistema <> "*" then 
                                                    if ClaseApp = 1 then response.write "selected"
                                                else
                                                    response.write "selected"
                                                end if
                                            %> >Web App</option>
                        </select> 
                    </div>

                    <div class="line">
                        <label class="label normal">Ordenamiento</label>
                        <input class="field small"  id="IndiceOrdenamiento" name="IndiceOrdenamiento" type="text" required <% if Sistema <> "*" then response.write "value='" & IndiceOrdenamiento & "'" %> />
                    </div>

                    <div class="line">
                        <label class="label normal">Nombre Icono</label>
                        <input class="field normal" id="Icono" name="Icono" type="text" required <% if Sistema <> "*" then response.write "value='" & Icono & "'" %> />
                    </div>

                    <div class="line">
                        <label class="label normal">Bitacora</label>

                        <select class="field small" name="sBitacora" id="sBitacora" >
                            <option value="0" <% if Sistema <> "*" then 
                                                    if sBitacora = 0 then response.write "selected"
                                                end if
                                            %> >Sin Bitácora</option>

                            <option value="1" <% if Sistema <> "*" then 
                                                    if sBitacora = 1 then response.write "selected"
                                                else
                                                    response.write "selected"
                                                end if
                                            %> >Con Bitácora</option>
                            </select>
                    </div>

                    <% if TieneVariables(Sistema) > 0 then %>
                        <div class="line">
                            <label class="label normal">Variables:</label>
                            <label class="label full section"> 
                                <% VariablesAsignadas(Sistema) %>
                            </label>
                        </div>
                        
                    <% end if %>    

                    <br />                             
                </div>
            </form>             
        </div>

        <br />

        <script type="text/javascript">
            function volver() {
                var vinculo = "lista.asp?op=<%= ordenadoPor %>";
                window.location.href = vinculo;
            }

            function reColor(Campo) {
                const elemento = document.getElementById(Campo);
                const color = elemento.value;
                
                if (!/^#[0-9A-Fa-f]{6}$/.test(color)) {
                    console.warn(`El valor "${color}" no es un color hexadecimal válido`);
                    return;
                }

                elemento.style.backgroundColor = color;
            }

            function grabar() {
                document.getElementById("formulario").submit();
            }           
        </script> 

        <% Con.close: set Con = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->            
    </body>
</html>