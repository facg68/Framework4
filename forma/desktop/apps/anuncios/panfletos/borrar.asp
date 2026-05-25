<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function NombreObjeto(secuencia)
                dim c, t, sqlStr

                sqlStr = "SELECT Objeto FROM seg_Panfletos " & _
                          "WHERE (Secuencia = " & Secuencia & ");"

                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set t = c.execute(sqlStr)
                        NombreObjeto = trim(t("Objeto"))
                    t.close: set t = nothing
                c.close: set c = nothing      
            end function

            sub BorrarObjeto(secuencia)
                On Error resume next

                    dim imagen, path, objeto

                    Objeto = NombreObjeto(Secuencia)

                    if len(trim(Objeto)) > 0 then
                        path = Server.MapPath("\") & "\desktop\apps\anuncios\pdf\" 
                        imagen = path & Objeto

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
            dim cc, tt, sqlString, sec
            dim verPantalla, estatusAnuncio, ordenadoPor

            estatusAnuncio = request.querystring("e")
            ordenadoPor = request.querystring("op")                 
            sec = Request.QueryString("p")

            sqlString = "DELETE FROM seg_Panfletos " & _
                        "WHERE (Secuencia = " & sec & ");"

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")
                BorrarObjeto sec            
                cc.execute(sqlString)
            cc.close: set cc = nothing

            response.redirect "lista.asp?e=" & estatusAnuncio & "&op=" & ordenadoPor
        %>    
    </body>
</html>