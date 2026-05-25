<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <style>
            body {
                margin: 0;
                padding: 0;
                overflow: hidden;
                background-color: black;
            }             
        </style>
            
        <%        
            function AnuncioActual(UltimoAnuncio)
                dim cc, tt, ssql

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")

                AnuncioActual = 0

                if UltimoAnuncio <> "" then
                    ssql = "SELECT TOP (1) Orden " & _
                            "FROM dbo.seg_Lista_Anuncios_TVs " & _
                            "WHERE (CodigoPantalla = '" & Request.Cookies("pantalla") & "') " & _
                            "AND (Orden > " & UltimoAnuncio & ") " & _
                        "ORDER BY Orden;"

                    set tt = cc.execute(ssql)

                    if not (tt.bof or tt.eof) then
                        AnuncioActual = tt("Orden")
                    else
                        tt.close: set tt = nothing

                        ssql = "SELECT Top 1 Orden FROM dbo.seg_Lista_Anuncios_TVs " & _
                               "WHERE (CodigoPantalla = '" & Request.Cookies("pantalla") & "') " & _
                               "ORDER BY Orden;"

                        set tt = cc.execute(ssql)

                        if not (tt.bof or tt.eof) then
                            AnuncioActual = tt("Orden")
                        end if
                    end if
                else
                    ssql = "SELECT Top 1 Orden FROM dbo.seg_Lista_Anuncios_TVs " & _
                           "WHERE (CodigoPantalla = '" & Request.Cookies("pantalla") & "') " & _
                           "ORDER BY Orden;"

                    set tt = cc.execute(ssql)

                    if not (tt.bof or tt.eof) then
                        AnuncioActual = tt("Orden")
                    end if            
                end if

                tt.close: set tt = nothing
                cc.close: set cc = nothing                
            end function   

            function HayAnuncios()
                dim cc, tt, ssql, Cookie_Pantalla

                HayAnuncios = False
                VerificarPantallaCookie

                Cookie_Pantalla = Request.Cookies("Pantalla")

                if Len(Trim(Request.Cookies("pantalla"))) = 0 then
                    Response.Cookies("Pantalla") = ParametroUsuario("anuncios", "anuncios_display")
                    Response.Cookies("Pantalla").Expires = Date() + 3000    
                end if

                ssql = "SELECT COUNT(*) AS Cuantos FROM (SELECT Secuencia FROM dbo.seg_Lista_Anuncios_TVs " & _
                        "WHERE (CodigoPantalla = '" & Request.Cookies("pantalla") & "')) AS t"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.execute(ssql)
                    
                    if tt("Cuantos") > 0 then
                        HayAnuncios = True
                    end if

                tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function

            Function ParametroUsuario(Sistema, Parametro)
                dim seg_conn, t, cmdString, sis, proc

                cmdString = "SELECT Valor " & _
                            "FROM seg_Usuarios_Parametros " & _
                            "WHERE (Sistema = '" & Sistema & "') " & _
                            "AND (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                            "AND (Parametro = '" & Parametro & "');"

                set seg_conn = Server.CreateObject("ADODB.Connection")
                seg_conn.open Application("Conn") 
                    set t = seg_conn.execute(cmdString)
                        if not (t.bof or t.eof) then
                            ParametroUsuario = t("Valor")
                        else
                            ParametroUsuario = NULL
                        end if   
                    t.close: set t = nothing
                seg_conn.close: set seg_conn = nothing               
            end Function      

            Sub VerificarPantallaCookie()
                On Error Resume Next
                dim v

                v = Request.Cookies("Pantalla")

                If Err.Number <> 0 Then
                    On Error Goto 0 
                    Response.Cookies("Pantalla") = ""
                End If
                On Error Goto 0 
            end Sub      
        %>

       <script>
            function volver() {
                var vinculo = "lista.asp";
                window.location.href = vinculo;          
            }   

            window.addEventListener("click", volver);            
        </script>        
    </head>

        <%
            dim con, t, sqlString, anuncio
            dim segundos, Orden

            if HayAnuncios() then
                anuncio = AnuncioActual(Request.QueryString("l"))

                if (len(Trim(anuncio)) > 0) then
                    sqlString = "SELECT * FROM dbo.seg_Lista_Anuncios_TVs " & _
                                "WHERE (CodigoPantalla = '" & Request.Cookies("pantalla") & "') " & _
                                "AND (Orden = " & Anuncio & "); "
                else
                    sqlString = "SELECT Top 1 FROM dbo.seg_Lista_Anuncios_TVs " & _
                                "WHERE (CodigoPantalla = '" & Request.Cookies("pantalla") & "') " & _
                                "ORDER BY Orden;"                
                end if

                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")
                set t = con.execute(sqlString)
                
                    Orden = t("Orden")
                    Segundos = t("Segundos")

                    select case t("Tipo")               
                        case 1 ' Imagen
                            response.write "<body style='background-color: rgb(0,0,0);'>"
                                if t("EscalarImagen") = 1 then 
                                    ' Ampliar al 100%
                                    response.write "<img src='/forma/desktop/apps/anuncios/publicaciones/" & t("NombreImagen") & "' width='100%' height='100%'>"
                                else 
                                    ' Centrar
                                    response.write "<div style='width: 95%; margin: auto;'>"
                                        response.write "<img src='/forma/desktop/apps/anuncios/publicaciones/" & t("NombreImagen") & "' height='100%'>"
                                    response.write "</div>"
                                end if
                            
                        case 2 ' HTML... 
                            response.write "<body style='background-color: rgb(255, 255, 255);'>"                    
                                response.write t("Cuerpo")

                        case 3 ' Video... 
                            response.write "<body style='background-color: rgb(0, 0, 0);'>"                    
                                response.write "<video autoplay loop style='" 
                                    if t("EscalarImagen") = 1 then 
                                        ' Ampliar al 100%
                                        response.write "width: 100vw; height: 100vh; object-fit: fill; background-color: black;"
                                    else 
                                        ' Centrar
                                        response.write "width: 100%;height: 100%; object-fit: contain; background-color: black;"
                                    end if                                        
                                response.write "'>"
                                    response.write "<source src='/forma/desktop/apps/anuncios/publicaciones/" & t("NombreImagen") & "' type='video/mp4'>"
                                response.write "</video>"                                   
                                
                        case 4 ' Pagina Web (interna)
                            response.write "<body style='background-color: rgb(0, 0, 0);'>"                        
                                response.write "<iframe id='dispFrame' width='100%' height='98%' src='" & t("Cuerpo") & "' frameborder='0' allowFullScreen='true'>"
                                response.write "</iframe>"     

                        case 5 ' Visor de PowerBI
                            response.write "<body style='background-color: rgb(0, 0, 0);'>"                        
                                response.write "<iframe id='dispFrame' width='100%' height='98%' "
                                response.write "src='https://app.powerbi.com/view?r=" & t("Cuerpo") & "' "
                                response.write "frameborder='0' allowFullScreen='true'>"
                                response.write "</iframe>"                                                                 
                    end select

                t.close: set t = nothing    
                con.close: set con = nothing                       
            else
                response.write "No Hay Anuncios para esta TV en este momento"
                Orden = 0
                Segundos = 120               
            end if
        %>

        <script>
            setTimeout(function(){
                window.location.href = "display.asp?l=<%= Orden %>";
            }, <%= (Segundos * 1000 ) %>);    
        </script> 
    </body>
</html>