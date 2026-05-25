<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">
        <title>Editar Publicación</title>       
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  

        <%
            thisSystem = "anuncios"
            thisProcess = "anu.0100"
            SysLockOut

            function fechaFormulario(fechaServer)
                dim d, m, a, h, min

                d = right("00" & DAY(fechaServer), 2)
                m = right("00" & MONTH(fechaServer), 2)
                a = YEAR(fechaServer)

                h = right("00" & HOUR(fechaServer), 2)
                min = right("00" & MINUTE(fechaServer), 2)

                fechaFormulario = d & "/" & m & "/" & a & " " & h & ":" & min
            end function  

            function PantallaSeleccionada(anuncio, pantalla)
                dim cc, tt

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn") 
                    set tt = cc.execute("SELECT * " & _
                                        "FROM seg_Anuncios_Asignaciones " & _
                                        "WHERE (Anuncio = " & anuncio & ") " & _
                                        "AND (Pantalla = '" & pantalla & "');")
                        if not (tt.bof or tt.eof) then
                            PantallaSeleccionada = True
                        else
                            PantallaSeleccionada = False
                        end if
                    tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function            
        %> 
        
        <style>
            p {
                margin-bottom: 10px;
                line-height: 1.5;
            }          
        </style>        
    </head>

    <body plantilla="normal" reserva="150">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->   

        <%
            dim con, t, p, sqlString, codAnuncio, verPantalla, estatusAnuncio, ordenadoPor, noPantallas

            codAnuncio = request.querystring("a")
            verPantalla = request.querystring("tv")
            estatusAnuncio = request.querystring("e")
            ordenadoPor = request.querystring("op")  
            republicar = request.querystring("r")  
            origen = request.querystring("w")               

            if republicar = "" then republicar = 0

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            if codAnuncio <> "*" then
                set t = con.execute("SELECT * FROM seg_Anuncios WHERE Secuencia = " & codAnuncio & ";")
            end if
        %>

        <br />

        <div style="width: 100%; margin: auto;">
            <form id="formulario"  name="formulario" method="post" action="grabar.asp">
                <div class="no-ver">
                    <%
                        cCodigo = "*"
                        cNombreImagen = "*"

                        if codAnuncio <> "*" then 
                            cCodigo = t("Codigo")
                            cNombreImagen = t("nombreImagen")
                        end if
                    %>

                    <input id="secuencia"       name="secuencia"        type="text" value="<%= codAnuncio %>" >
                    <input id="propietario"     name="propietario"      type="text" value="<%= Request.Cookies("usuario") %>" >
                    <input id="verPantalla"     name="verPantalla"      type="text" value="<%= verPantalla %>" >
                    <input id="estatusAnuncio"  name="estatusAnuncio"   type="text" value="<%= estatusAnuncio %>" >
                    <input id="ordenadoPor"     name="ordenadoPor"      type="text" value="<%= ordenadoPor %>" >
                    <input id="republicar"      name="republicar"       type="text" value="<%= republicar %>" >
                    <input id="origen"          name="origen"           type="text" value="<%= origen %>" >
                    <input id="codigo"          name="codigo"           type="text" value="<%= cCodigo %>" >
                    <input id="nombreImagen"    name="nombreImagen"     type="text" value="<%= cNombreImagen %>" >
                </div>    

                <table style="width: 95%; margin: auto;"> 
                    <tr>
                        <td style="width: 55%; font-size: 24px;">
                            <h3>
                                <%
                                    if codAnuncio = "*" then
                                        response.write "Nueva Publicación"
                                    else
                                        if republicar = "0" then
                                            response.write "Editar Publicación"
                                        else
                                            response.write "Volver a Publicar (Copia)"
                                        end if
                                    end if
                                %>
                            </h3>
                        </td>

                        <td style="width: 45%; text-align: right;">
                            <button class="form-btn verde normal" type="submit">Grabar</button>       
                            
                            <a href="lista.asp?tv=<%= verPantalla %>&e=<%= estatusAnuncio %>&op=<%= ordenadoPor %>">
                                <button class="form-btn rojo normal" type='button'>Cancelar</button>
                            </a>                    
                        </td>
                    </tr>
                </table>

                <div class="main main-scroll">
                    <div class="line">
                        <label class="label normal">Titulo</label>
                        <input class="field xl" id="titulo" name="titulo" type="text" required <% if codAnuncio <> "*" then response.write "value='" & t("Titulo") & "'" %> />
                    </div>

                    <div class="line">
                        <label class="label normal">Desde</label>
                        <input class="field small" id="inicio" name="inicio" type="text" placeholder="dd/mm/aaaa hh:mm" required <% if codAnuncio <> "*" then response.write "value='" & fechaFormulario(t("Inicio")) & "'" %> />
                    </div>

                    <div class="line">
                        <label class="label normal">Hasta</label>
                        <input class="field small" id="fin" name="fin" type="text" placeholder="dd/mm/aaaa hh:mm" required <% if codAnuncio <> "*" then response.write "value='" & fechaFormulario(t("Fin")) & "'" %> />
                    </div>                    

                    <div class="line">
                        <label class="label normal">Tipo de Publicación</label>
                        <select class="field normal" name="tipo" id="tipo" onChange="toggleOperacion();">
                            <option value="1" <% if codAnuncio <> "*" then 
                                                    if t("Tipo") = 1 then response.write "selected"
                                                else
                                                    response.write "selected"
                                                end if
                                            %> >Imagen</option>
                            <option value="3" <% if codAnuncio <> "*" then 
                                                    if t("Tipo") = 3 then response.write "selected"
                                                end if
                                            %> >Video</option>
                            <option value="2" <% if codAnuncio <> "*" then 
                                                    if t("Tipo") = 2 then response.write "selected"
                                                end if
                                            %> >HTML</option>
                            <option value="4" <% if codAnuncio <> "*" then 
                                                    if t("Tipo") = 4 then response.write "selected"
                                                end if
                                            %> >Pagina Web</option>  
                            <option value="5" <% if codAnuncio <> "*" then 
                                                    if t("Tipo") = 5 then response.write "selected"
                                                end if
                                            %> >Visor de PowerBI</option>                                                                                       
                        </select> 
                    </div>

                    <div class="line" id="imagen2">
                        <label class="label normal">Manejo de Imagen</label>
                        <select class="field normal" name="EscalarImagen">
                            <option value="1" <% if codAnuncio <> "*" then 
                                                    if t("EscalarImagen") = 1 then response.write "selected"
                                                else
                                                    response.write "selected"
                                                end if
                                            %>>Escalar</option>
                            <option value="2" <% if codAnuncio <> "*" then 
                                                    if t("EscalarImagen") = 2 then response.write "selected"
                                                end if
                                            %>>Centrar</option>
                        </select>
                    </div>  

                    <div class="line" id="htmle" style="display: <%
                        if codAnuncio <> "*" then
                            if t("Tipo") = 2 or t("Tipo") = 4 or t("Tipo")  = 5 then
                                response.write "block"
                            else
                                response.write "none"
                            end if
                        else
                            response.write "none"
                        end if
                     %>;">
                        <label class="label normal">Codigo HTML</label>
                        <textarea class="field" name="Cuerpo" id="Cuerpo" 
                                    rows=10 cols=80 class="vbControl_Verde item" 
                                    style="border: 1px solid rgb(215, 215, 215); 
                                        padding: 10px; 
                                        font-family: courier; 
                                        font-size: 16px; 
                                        width: 100%;"><% 
                                    if codAnuncio <> "*" then 
                                        response.write t("Cuerpo")
                                    end if
                                    %></textarea>
                    </div>  

                    <div class="line" id="htmle2" style="display: none;">
                        <label class="label normal">Ver Código HTML Como</label>
                        <select class="field normal"name="VerComo" id="VerComo">
                            <option value="0" <% if codAnuncio <> "*" then 
                                                    if t("VerComo") = 1 then response.write "selected"
                                                else
                                                    response.write "selected"
                                                end if
                                            %>>Cuadro de Texto</option>
                            <option value="1" <% if codAnuncio <> "*" then 
                                                    if t("VerComo") = 1 then response.write "selected"
                                                end if
                                            %>>Editor Extendido</option>
                        </select>
                    </div>             

                    <div class="line">
                        <label class="label normal">Segundos en Pantalla</label>
                        <input class="field small"
                                id="Segundos" name="Segundos" 
                                type="text" placeholder="000" required 
                                <% if codAnuncio <> "*" then response.write "value='" & t("Segundos") & "'" %> 
                        />
                    </div>                            

                    <div class="line">
                        <label class="label normal">Publicar En</label>  
                        <label class="label section full">                
                            <%
                                sqlString = "SELECT up.Pantalla, p.Nombre " & _
                                                "FROM dbo.seg_Anuncios_Pantallas_Usuarios AS up " & _
                                        "INNER JOIN dbo.seg_Anuncios_Pantallas AS p " & _ 
                                                "ON up.Pantalla = p.Pantalla " & _
                                                "WHERE (up.Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                            "ORDER BY p.Nombre;"

                                set p = con.execute(sqlString)

                                if not (p.bof or p.eof) then
                                    noPantallas = 0

                                    Do

                                        nombreCampo = "pantalla_" & Trim(p("Pantalla"))

                                        response.write "<p>"
                                            response.write "<input type='checkbox' id='" & nombreCampo & "' name='" & nombreCampo & "' value='1' " 
                                                if codAnuncio <> "*" then 
                                                    if PantallaSeleccionada(codAnuncio, p("pantalla")) = true then
                                                        response.write "checked"
                                                    end if
                                                end if                                
                                            response.write ">"
                                            
                                            response.write "<label>&nbsp;&nbsp;" & p("Nombre") & "</label>"
                                        response.write "</p>"

                                    p.MoveNext
                                    Loop Until (p.eof)
                                else
                                    response.write "El usuario actual no tiene permiso para publicar en ninguna pantalla."
                                    noPantallas = 1
                                end if

                                p.close: set p = nothing
                            %>
                        </div>
                    </div> 
                </div>  
            </form>
        </div>

        <script>
            function toggleOperacion() {
                var elemento_a_evaluar = document.getElementById("tipo").value;

                var elemento2 = document.getElementById("imagen2");
                var elemento3 = document.getElementById("htmle");

                if (elemento_a_evaluar == "2" || elemento_a_evaluar == "4" || elemento_a_evaluar == "5") {
                    elemento2.style.display = "none";          
                    elemento3.style.display = "block";          
                } else {
                    elemento2.style.display = "block";
                    elemento3.style.display = "none";
                }
            }	    
            
            mask(document.getElementById('inicio'), ['99/99/9999 99:99']);
            mask(document.getElementById('fin'),    ['99/99/9999 99:99']);
        </script> 

        <%
            if codAnuncio <> "*" then
                t.close: set t = nothing
            end if    

            con.close: set con = nothing
        %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->          
    </body>
</html>