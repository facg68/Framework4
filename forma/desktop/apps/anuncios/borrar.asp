<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function NombreImagen(secuencia)
                dim c, t, sqlStr

                sqlStr = "SELECT NombreImagen FROM seg_Anuncios " & _
                          "WHERE (Secuencia = " & Secuencia & ");"

                response.write "fx NombreImagen - Secuencia: " & Secuencia & "<br/>"
                response.write "fx NombreImagen - sqlStr: " & sqlStr & "<br/>"

                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                set t = c.execute(sqlStr)

                    NombreImagen = trim(t("NombreImagen"))

                    response.write "fx NombreImagen - NombreImagen: " & NombreImagen & "<br/>"

                t.close: set t = nothing
                c.close: set c = nothing      
            end function

            sub BorrarImagen(secuencia)
                On Error resume next

                    dim imagen, path, objeto

                    Objeto = NombreImagen(Secuencia)

                    response.write "Secuencia: " & Secuencia & "<br/>"
                    response.write "Objeto: " & Objeto & "<br/>"

                    if len(trim(Objeto)) > 0 then
                        path = Server.MapPath("\") & "\imagenes\anuncios\" 
                        imagen = path & Objeto

                        response.write "Path: " & path & "<br/>"
                        response.write "Imagen: " & Imagen & "<br/>"

                        Set fso = CreateObject("Scripting.FileSystemObject")
                            fso.DeleteFile imagen
                        set fso = nothing      
                    end if   

                On Error Goto 0      
            end sub
        %>
    </head>

    <body>
        <%
            dim cc, tt, sqlString, sec, origen
            dim verPantalla, estatusAnuncio, ordenadoPor

            verPantalla = request.querystring("tv")
            estatusAnuncio = request.querystring("e")
            ordenadoPor = request.querystring("op")                 
            origen = request.querystring("w")                 

            sec = Request.QueryString("a")

            sqlString = "DELETE FROM seg_Anuncios " & _
                        "WHERE (Secuencia = " & sec & ");"

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")
                borrarImagen sec            
                cc.execute(sqlString)
            cc.close: set cc = nothing

            if origen = 1 then
                response.redirect "lista_total.asp?tv=" & verPantalla & "&e=" & estatusAnuncio & "&op=" & ordenadoPor
            else
                response.redirect "lista.asp?tv=" & verPantalla & "&e=" & estatusAnuncio & "&op=" & ordenadoPor
            end if
        %>    
    </body>
</html>