<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Subir Objeto</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->
        <%
            thisSystem = "anuncios"
            thisProcess = "anu.0100"
            SysLockOut
        %>

        <style>
            table, tr, td, th, tbody {
                width: 100%;
            }

            td, th {
                padding: 5px;
                font-size: 16px;
            }   

            .top {
                background-color: rgb(71,71,71);
                color: white;
            }             
        </style>   

        <%    
            sub append(byRef Cadena, NuevaCadena)
                if NuevaCadena <> "" then
                    Cadena = Cadena & NuevaCadena
                end if
            end sub

            function tipoArchivo(anuncio)
                dim cc, tt, ssql

                ssql = "SELECT Tipo FROM seg_Anuncios WHERE Secuencia = " & anuncio & ";"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.Open Application("Conn")
                set tt = cc.execute(ssql)

                if not (tt.bof or tt.eof) then
                    select case tt("tipo")
                        case 1
                            tipoArchivo = "image/*"
                        case 2
                            tipoArchivo = NULL
                        case 3
                            tipoArchivo = "video/*"                        
                    end select                    
                else
                    tipoArchivo = NULL
                end if

                tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function

            function Tipo(anuncio)
                dim cc, tt, ssql

                ssql = "SELECT Tipo FROM seg_Anuncios WHERE Secuencia = " & anuncio & ";"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.Open Application("Conn")
                set tt = cc.execute(ssql)

                Tipo = tt("tipo")

                tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function   

            function TituloAnuncio(anuncio)    
                dim cc, tt, ssql

                ssql = "SELECT Titulo FROM seg_Anuncios WHERE Secuencia = " & anuncio & ";"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.Open Application("Conn")
                set tt = cc.execute(ssql)

                TituloAnuncio = tt("Titulo")

                tt.close: set tt = nothing
                cc.close: set cc = nothing                 
            end function
        %>  
    </head>

    <body plantilla="normal" reserva="175">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    
        <%
            dim con, t, sqlString, anuncio, usuario, nombreimagen
            dim verPantalla, estatusAnuncio, ordenadoPor, origen

            usuario = Request.Cookies("usuario")
            origen = Request.QueryString("w")

            response.cookies("edit_anuncio") = Request.QueryString("a")
            response.cookies("edit_verPantalla") = Request.QueryString("tv")
            response.cookies("edit_estatusAnuncio") = Request.QueryString("e")
            response.cookies("edit_ordenadoPor") = Request.QueryString("op")
            response.cookies("edit_origen") = Request.QueryString("w")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            sqlString = "SELECT NombreImagen " & _
                        "FROM seg_Anuncios " & _
                        "WHERE (Propietario = '" & usuario & "') " & _
                        "AND (Secuencia = '" & Request.Cookies("edit_anuncio") & "');"

            set t = con.execute(sqlString)
                nombreimagen = t("nombreimagen")
            t.close: set t = nothing
        %>  

        <br />        

        <form id="frm_adjuntos" name="frm_adjuntos" 
              style="width:98%; margin: auto;"
              action="subir_objeto2.asp" method="post" 
              enctype="multipart/form-data"> 

            <table style="width: 95%; margin: auto;">
                <tr>
                    <td style="width: 55%; text-align: left; font-size: 18px;">
                        <span style="font-size: 24px;">
                            <%= TituloAnuncio(request.cookies("edit_anuncio")) %><br />
                            <span style="font-size: 18px;">
                                OBJETO <%= nombreimagen %>
                            </span>
                        </span>
                    </td>

                    <td style="width: 45%; text-align: right;">
                        <button class="form-btn rojo normal" onclick="Volver()" type="button">Cancelar</button>                                                
                        <button class="form-btn verde normal" type="submit">Actualizar</button>
                    </td>
                </tr>
            </table>

            <div class="main main-scroll">
                <div class="line">      
                    <table>
                        <tr>
                            <td style="text-align:center; width:25%;">
                                <% 
                                    fotoPath = "/forma/desktop/apps/anuncios/publicaciones/" & nombreimagen 

                                    select case Tipo(Request.Cookies("edit_anuncio")) 
                                        case 1
                                            '
                                            ' Imagen
                                            '
                                            %>
                                                <img src="<%= fotoPath %>" 
                                                    onerror="this.src='/core/imagenes/misc/foto.jpg'" 
                                                    style="width:100%; height:auto;">
                                            <%
                                        case 2
                                            '
                                            ' HTML
                                            '
                                            response.write "<img src='/imagenes/misc/html.jpg' style='width:100%; height:auto;'>"

                                        case 3
                                            '
                                            ' Video
                                            '
                                            if (len(trim(nombreimagen)) = 0) OR (ISNULL(nombreimagen)) then
                                                response.write "<img src='/imagenes/misc/video.jpg' style='width:100%; height:auto;'>"
                                            else
                                                response.write "<video width='300px' autoplay>"
                                                    response.write "<source src='" & fotoPath & "' type='video/mp4'>"
                                                response.write "</video>"                                                        
                                            end if
                                    end select
                                %>
                            </td>

                            <td style="width:5%;">&nbsp;</td>

                            <td style="width:65%;">
                                <input class="field xxl" type="file" id="File" name="File" accept="<%= tipoArchivo(Request.Cookies("edit_anuncio")) %>" /> 
                                <input class="no-ver" type="text" id="anuncio" name="anuncio" value="<%= Request.Cookies("edit_anuncio") %>" /> 
                            </td>

                            <td style="width:5%;">&nbsp;</td>
                        </tr>
                    </table>
                </div> 
            </div>
        </form> 

        <script>
            function Volver(anuncio) {
                var vinculo = "<%
                        if origen = 1 then
                            response.write "lista_total"
                        else
                            response.write "lista" 
                        end if
                    %>.asp?tv=<%= Request.Cookies("edit_verPantalla") %>&e=<%= Request.Cookies("edit_estatusAnuncio") %>&op=<%= Request.Cookies("edit_ordenadoPor")%>";
                window.location.href = vinculo;
            }          
        </script>      

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>


