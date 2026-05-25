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
            dim con, t, sqlString, anuncio, verPantalla, estatusAnuncio, ordenadoPor, origen

            verPantalla = request.querystring("tv")
            estatusAnuncio = request.querystring("e")
            ordenadoPor = request.querystring("op")           
            origen = request.querystring("w")           
        %>
    
        <script>
            function volver() {
                var vinculo = "<%
                        if origen = 1 then
                            response.write "lista_total"
                        else
                            response.write "lista" 
                        end if
                    %>.asp?tv=<%= Request.Cookies("edit_verPantalla") %>&e=<%= Request.Cookies("edit_estatusAnuncio") %>&op=<%= Request.Cookies("edit_ordenadoPor")%>";
                window.location.href = vinculo;          
            }   

            window.addEventListener("click", volver);            
        </script>
    </head>

    <%
        set con = Server.CreateObject("ADODB.Connection")
        con.open Application("Conn")

        sqlString = "SELECT * FROM dbo.seg_Anuncios " & _
                    "WHERE (secuencia = " & Request.QueryString("a") & ");"            

        set t = con.execute(sqlString)

        select case t("Tipo")               
            case 1 ' Imagen
                response.write "<body style='background-color: rgb(0,0,0);'>"
                    if t("EscalarImagen") = 1 then 
                        ' Ampliar al 100%
                        response.write "<img src='/forma/desktop/apps/anuncios/publicaciones/" & t("NombreImagen") & "' "
                        response.write "style='width: 100%; height: 100%;'>"
                    else 
                        ' Centrar
                        response.write "<div style='width: 95%; margin: auto;'>"
                            response.write "<img src='/forma/desktop/apps/anuncios/publicaciones/" & t("NombreImagen") & "' "
                            response.write "style='height: 100%;'>"
                        response.write "</div>"
                    end if
                response.write "</body>"
                
            case 2 ' HTML... 
                response.write "<body style='background-color: rgb(255, 255, 255);'>"                    
                    response.write t("Cuerpo")
                response.write "</body>"

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
                    response.write "<iframe id='dispFrame' width='100%' height='100%' src='" & t("Cuerpo") & "' frameborder='0' allowFullScreen='true'>"
                    response.write "</iframe>"     

            case 5 ' Visor de PowerBI
                response.write "<body style='background-color: rgb(0, 0, 0);'>"                        
                    response.write "<iframe id='dispFrame' width='100%' height='100%' "
                    response.write "src='https://app.powerbi.com/view?r=" & t("Cuerpo") & "' "
                    response.write "frameborder='0' allowFullScreen='true'>"
                    response.write "</iframe>"                   
        end select

        response.write "</body>"   

        t.close: set t = nothing    
        con.close: set con = nothing    
    %>
</html>