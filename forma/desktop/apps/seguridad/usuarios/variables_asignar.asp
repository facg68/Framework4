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
            
            .vbControl_B_Enabled {
                background-color: rgb(240, 255, 227);
                color: rgb(0, 0, 0);
                padding: 5px;
                border: 1px solid rgb(28, 69, 117);
            } 
        </style>

        <%
            dim pCon, ptt, pProc, sqlString, Sistema, Parametro, ordenadoPor, tt
            dim Nombre, Descripcion, TipoRol, cuantosSistemas, primerTab

            Sistema = request.querystring("s")
            Parametro = request.querystring("p")
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

            Function ConsultaSQL(Sistema, Parametro) 
                dim sqlString

                sqlString = "SELECT ConsultaSQL AS Consulta FROM seg_Parametros WHERE (Sistema = '" & Sistema & "' AND Parametro = '" & Parametro & "');"

                set ptt = pCon.execute(sqlString)   
                    ConsultaSQL = ptt("Consulta") 
                    ConsultaSQL = Replace(ConsultaSQL, "|", "'")                    
                ptt.close: set ptt = nothing                
            End Function

            function CuantosUsuariosQuedan(Sistema, Parametro)
                sqlString = "exec dbo.seg_Params_UsuariosNoAsignados '" & Sistema & "', '" & Parametro & "'"
                CuantosUsuariosQuedan = 0

                set ptt = pCon.execute(sqlString)   
                    if not (ptt.bof or ptt.eof) then
                        do
                            CuantosUsuariosQuedan = CuantosUsuariosQuedan + 1
                            ptt.MoveNext
                        loop until ptt.eof
                    end if
                ptt.close: set ptt = nothing
            end function

            sub UsuariosAsignados(Sistema, Parametro)
                %>
                    <table style="width: 100%;" style="padding: 0px;"> 
                        <tr style="background-color: rgb(0, 0, 0); color: rgb(255, 255, 255);">
                            <td class="top borde" style="width:  5%; text-align: center; padding: 5px;">&nbsp;</td>
                            <td class="top borde" style="width: 25%; text-align: left;   padding: 5px;">Usuario</td>
                            <td class="top borde" style="width: 25%; text-align: left;   padding: 5px;">Cargo</td>
                            <td class="top borde" style="width: 20%; text-align: left;   padding: 5px;">Variable</td>
                            <td class="top borde" style="width: 20%; text-align: left;   padding: 5px;">Valor</td>
                            <td class="top borde" style="width:  5%; text-align: center; padding: 5px;">&nbsp;</td>              
                        </tr>

                        <% if CuantosUsuariosQuedan(Sistema, Parametro) > 0 then %>
                            <tr>
                                <td class="borde" style="width: 5%; text-align: center; font-size: 30px; font-weight: bold;">+</td>              

                                <td colspan="4" class="borde" style="width: 55%;">
                                    <%
                                        sqlString = "exec dbo.seg_Params_UsuariosNoAsignados '" & Sistema & "', '" & Parametro & "'"

                                        set tt = pCon.execute(sqlString)
                                            response.write "<select class='field full' name='usuNuevo' id='usuNuevo'>"
                                                response.write "<option value='*'> - - Seleccione un Usuario - - </option>"

                                                Do
                                                    response.write "<option value='" & tt("usuario") & "'>" & tt("usuNombre") & "</option>"
                                                    tt.MoveNext
                                                Loop Until tt.eof

                                            response.write "/<select>"
                                        tt.close: set tt = nothing
                                    %>                                    
                                </td>              

                                <td class="borde" style="width: 5%; text-align: center; background-color: rgb(240, 255, 227);">
                                    <button type="button" class="form-btn verde" onclick="NuevaLinea()"><i class="fa fa-save fa-xl"></i></button>
                                </td>              
                            </tr>    
                        <% end if %>

                        <%
                            sqlString = "SELECT up.Usuario, up.Sistema, up.Parametro, p.TipoParametro, u.usuNombre, u.usuCargo, up.Valor " & _
                                          "FROM dbo.seg_Usuarios_Parametros AS up INNER JOIN dbo.seg_Usuarios AS u ON up.Usuario = u.usuCodigo " & _
                                    "INNER JOIN dbo.seg_Parametros AS p ON up.Parametro = p.Parametro AND up.Sistema = p.Sistema " & _
                                    "WHERE up.Sistema = '" & Sistema & "' AND up.Parametro = '" & Parametro & "' " & _
                                      "ORDER BY u.usuNombre;"

                            set cbox = pCon.execute(sqlString)
                                if not (cbox.bof or cbox.eof) then
                                    contador = 0

                                    Do        
                                        contador = contador + 1
                                        nomCampo = "pValor_" & contador

                                        response.write "<tr>"
                                            response.write "<td style='text-align: center; padding: 5px; width: 5%' class='borde'>"
                                                response.write contador
                                            response.write "</td>"      

                                            response.write "<td style='text-align: left; padding: 5px; width: 25%' class='borde'>"
                                                response.write cbox("usuNombre")
                                            response.write "</td>"                                   

                                            response.write "<td style='text-align: left; padding: 5px; width: 25%' class='borde'>"
                                                response.write cbox("usuCargo")
                                            response.write "</td>"

                                            response.write "<td style='text-align: left; padding: 5px; width: 20%' class='borde'>"
                                                response.write cbox("Parametro")
                                            response.write "</td>"

                                            response.write "<td style='text-align: left; padding: 5px; width: 20%' class='borde'>"
                                                Select Case cbox("TipoParametro")
                                                    Case "1"
                                                        '
                                                        ' Es un Permiso... El Campo tiene un valor de 1 y está escondido
                                                        '
                                                        response.write "<input class='field full' id='" & nomCampo & " name=" & nomCampo & " type='text' "
                                                        response.write "value=" & cbox("Valor") & " style='visibility: hidden;' />"
                                                    Case "2"
                                                        '
                                                        ' El Campo es una Variable... Presentamos el input-box
                                                        '
                                                        %>
                                                            <input class="field full" id="<%= nomCampo %>" name="<%= nomCampo %>" type="text" value="<%= cbox("Valor") %>" 
                                                                onChange="actualizarValor('<%= nomCampo %>', '<%= cbox("Usuario") %>', '<%= cbox("Valor") %>')" />

                                                        <%
                                                    Case "3"
                                                        '
                                                        ' La Variable es un Campo "Si/No". Presentamos un Check-Box
                                                        '
                                                        %>
                                                            <input type="radio" id="<%= nomCampo & "1" %>" name="<%= nomCampo %>" value="1" <%
                                                                if cbox("Valor") = "1" then response.write " checked " %>
                                                                onChange="actualizarValor('<%= nomCampo & "1" %>', '<%= cbox("Usuario") %>', '<%= cbox("Valor") %>')">
                                                            &nbsp;&nbsp;
                                                            <label style="font-family: Verdana; font-size: 14px;" for="Si">Si</label>
                                                            
                                                            &nbsp;&nbsp;&nbsp;&nbsp;

                                                            <input type="radio" id="<%= nomCampo & "0" %>" name="<%= nomCampo %>" value="0" <%
                                                                if cbox("Valor") = "0" then response.write " checked " %>
                                                                onChange="actualizarValor('<%= nomCampo & "0" %>', '<%= cbox("Usuario") %>', '<%= cbox("Valor") %>')">
                                                            &nbsp;&nbsp;
                                                            <label style="font-family: Verdana; font-size: 14px;" for="No">No</label>
                                                        <%

                                                    Case "4"
                                                        '
                                                        ' Es una lista... Presentamos un Combo-Box
                                                        '
                                                        set lista = pCon.execute("SELECT * FROM seg_Parametros_Valores WHERE Sistema = '" & Sistema & "' AND Parametro = '" & Parametro & "' ORDER BY Descripcion;")

                                                        if not (lista.bof or lista.eof) then
                                                            %> <select class="field full" name="<%= nomCampo %>" id="<%= nomCampo %>" onChange="actualizarValor('<%= nomCampo %>', '<%= cbox("Usuario") %>', '<%= cbox("Valor") %>')"> <%
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
                                                        Consulta = ConsultaSQL(Sistema, Parametro)

                                                        set lista = pCon.execute(Consulta)

                                                        if not (lista.bof or lista.eof) then
                                                            %><select class="field full" name="<%= nomCampo %>" id="<%= nomCampo %>" onChange="actualizarValor('<%= nomCampo %>', '<%= cbox("Usuario") %>', '<%= cbox("Valor") %>')"> <%
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

                                                End Select

                                            response.write "</td>" 

                                            response.write "<td style='text-align: center; padding: 5px; width: 5%' class='borde'>"
                                                %><button class="form-btn rojo" type="button" onClick="BorrarPermiso('<%= cbox("Usuario") %>')"><%
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

    <body plantilla="normal" reserva="200">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->   

        <%
            set ptt = pCon.execute("SELECT * FROM seg_Parametros WHERE (Sistema = '" & Sistema & "') AND (Parametro = '" & Parametro & "');")
                Nombre = ptt("Parametro") 
                ValorDefault = "Valor por defecto: " & ptt("ValorDefault") 
                Descripcion = ptt("Descripcion") 
                Afectacion = ptt("Afectacion") 
            ptt.close: set ptt = nothing
        %>

        <div style="width: 95%; margin: auto;">
            <br />

            <table style="width: 100%;">
                <tr>
                    <td style="width: 70%;">
                        <%
                            response.write "<span style='font-size: 22px; font-weight: bold;'>" & Nombre & "</span><br />"
                            response.write "<span style='font-size: 16px'>" & ValorDefault & "</span><br />"
                            response.write "<span style='font-size: 16px'>" & Descripcion & "</span><br />"
                        %>
                    </td>

                    <td style="width: 30%; text-align: right;">
                        <button class='form-btn rojo normal'  type='button' onclick="volver()">Volver</button>&nbsp;&nbsp;
                    </td>
                </tr>
            </table>   
        </div>

        <br />

        <div style="width: 95%; margin: auto;">                                
            <form id="formulario"  name="formulario" method="post" action="grabar_rol.asp">
                <div class="no-ver">
                    <input id="ordenadoPor" name="ordenadoPor"  type="text" value="<%= ordenadoPor %>" />
                    <input id="Sistema"     name="Sistema"      type="text" value="<%= Sistema %>" />
                    <input id="Parametro"   name="Parametro"    type="text" value="<%= Parametro %>" />
                </div>

                <div class="main main-scroll">
                    <div class="line label-top">
                        <label class="label normal">Cómo Afecta</label>
                        <div class="label full section" style="background-color: rgb(235, 235, 235);">
                            <script src="/core/lib/tinymce/tinymce.min.js"></script>                        
                            
                            <textarea class="editor" id="Afectacion" name="Afectacion"> 
                                <%= Afectacion %>
                            </textarea>

                            <script>
                                tinymce.init({
                                    entity_encoding : "raw",
                                    selector: '.editor',
                                    license_key: 'gpl',
                                    height: 400,
                                    branding: false,
                                    promotion: false,                                    
                                    language: 'es',
                                    language_url: '/core/includes/es.js', 
                                    plugins: 'anchor autolink charmap codesample emoticons image link lists media searchreplace table visualblocks wordcount ',
                                    toolbar: 'undo redo | blocks fontfamily fontsize | bold italic underline strikethrough | link image media table mergetags | addcomment showcomments | spellcheckdialog a11ycheck typography | align lineheight | checklist numlist bullist indent outdent | emoticons charmap | removeformat'
                                });                              
                            </script>
                        </div>
                    </div>                                                

                    <div class="line label-top">
                        <label class="label normal">Usuarios</label>
                        <div class="label full section">
                            <% UsuariosAsignados Sistema, Parametro %> 
                        </div>
                    </div>
                </div>
            </form>
        </div>

        <script>
            function volver() {
                var vinculo = "variables.asp?s=<%= Sistema %>&o=<%= ordenadoPor %>";
                window.location.href = vinculo;
            }

            function NuevaLinea() {
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var usuNuevo = document.getElementById("usuNuevo").value;

                if (usuNuevo == "*") {
                    window.alert("Debe Seleccionar un Usuario");
                } else {
                    var vinculo ="variables_asignar_usuario.asp?o=" + ordenamiento + "&s=<%= Sistema %>&p=<%= Parametro %>&u=" + usuNuevo;
                    window.location.href = vinculo;    
                }
            }

            function actualizarValor(campo, usuario, valorActual) {
                var txtCampo = document.getElementById(campo);
                var nuevo_valor = document.getElementById(campo).value;
                var confirmacion = confirm("Desea actualizar el valor de la variable del usuario seleccionado?");
                var vinculo ="variables_actualizar.asp?s=<%= Sistema %>&p=<%= Parametro %>&u=" + usuario + "&v=" + nuevo_valor +  "&o=<%= ordenadoPor %>";     
                
                if (confirmacion) {     
                    window.location.href = vinculo;
                }
                else {
                    txtCampo.value = valorActual;
                };                
            }
     
            function BorrarPermiso(usuario) {
                var confirmacion = confirm("Desea quitar de este permiso al usuario seleccionado?");
                var vinculo = "variables_asignar_borrar.asp?o=<%= ordenadoPor %>&s=<%= Sistema %>&p=<%= Parametro %>&u=" + usuario;

                if (confirmacion) {     
                    window.location.href = vinculo;
                };
            }    
        </script> 

        <% pCon.close: set pCon = nothing %>
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>