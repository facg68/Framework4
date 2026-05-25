<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            sub BorrarImagenObjeto(Usuario, Objeto)
            On Error resume next
                dim imagen, imagen_s, path

                path = Server.MapPath("\") & "\perfiles\" & lcase(Usuario) & "\medios\"

                imagen = path & Objeto & ".jpg"
                imagen_s = path & Objeto & "_s.jpg"

                Set fso = CreateObject("Scripting.FileSystemObject")
                    fso.DeleteFile imagen
                    fso.DeleteFile imagen_s
                set fso = nothing               
            end sub

            sub BorrarImagenPaquete(Usuario, Paquete)
            On Error resume next
                dim imagen, imagen_s, path

                path = Server.MapPath("\") & "\perfiles\" & lcase(Usuario) & "\medios\"

                imagen = path & Paquete & ".jpg"
                imagen_s = path & Paquete & "_s.jpg"

                Set fso = CreateObject("Scripting.FileSystemObject")
                    fso.DeleteFile imagen
                    fso.DeleteFile imagen_s
                set fso = nothing               
            end sub            
        %>
    </head>

    <body>
        <%
            dim cc, tt, sqlString
            dim Usuario, Paquete, lista

            set cc = Server.CreateObject("ADODB.ConnectioN")
            cc.open Application("Conn")
                Usuario = Request.Cookies("Usuario")
                Paquete = Request.QueryString("p")

                set lista = cc.execute("SELECT Objeto FROM discos_Objetos WHERE Usuario = '" & Usuario & "' AND Paquete = '" & Paquete & "';")
                if not (lista.bof or lista.eof) then
                    Do
                        BorrarImagenObjeto Usuario, lista("Objeto")
                        lista.MoveNext
                    Loop Until lista.eof
                end if
                lista.close: set lista = nothing

                sqlString = "DELETE FROM discos_Paquetes " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Paquete = '" & Paquete & "');"
                
                cc.execute(sqlString)
            cc.close: set cc = nothing        

            BorrarImagenPaquete Usuario, Paquete

            response.redirect "lista.asp"
        %>
    </body>
</html>