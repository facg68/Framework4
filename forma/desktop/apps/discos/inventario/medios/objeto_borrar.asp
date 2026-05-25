<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            sub BorrarImagen(Usuario, Objeto)
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
        %>
    </head>

    <body>
        <%
            dim cc, tt, sqlString
            dim Usuario, Paquete, Objeto, Editor, nSecuencia

            set cc = Server.CreateObject("ADODB.ConnectioN")
            cc.open Application("Conn")
                Usuario = Request.Cookies("Usuario")

                Paquete = Request.QueryString("p")
                Objeto = Request.QueryString("o")
                Editor = Request.QueryString("e")

                sqlString = "DELETE FROM discos_Objetos " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Paquete = '" & Paquete & "') " & _
                            "AND (Objeto = '" & Objeto & "');"
                
                cc.execute(sqlString)
            cc.close: set cc = nothing        

            BorrarImagen Usuario, Objeto 

            response.redirect "editar.asp?m=" & Paquete
        %>
    </body>
</html>