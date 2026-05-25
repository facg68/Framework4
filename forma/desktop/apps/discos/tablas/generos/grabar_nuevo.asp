<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            dim con, t, tt, sqlString
            dim Usuario, Codigo, Nombre, Musica, Video, Juegos, Software, Libros, Hardware, Obsoleta

            function SecuenciaCodigo()
                dim cc, tt, sqlString

                sqlString = "SELECT ISNULL(MAX(CAST(Codigo AS numeric(10, 0))), 0) + 1 AS Valor " & _
                            "FROM dbo.discos_Tipos WHERE (Usuario = '" & Request.Cookies("Usuario") & "')"

                set cc = Server.CreateObject("ADODB.Connection")
                    cc.open Application("Conn")
                    set tt = cc.execute(sqlString)
                        SecuenciaCodigo = right( "00000000" & tt("Valor"), 8)
                    tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function
        %>
    </head>

    <body>
        <%
            Usuario = Request.Form("Usuario")
            VerTipo = Request.Form("Ver")
            Ordenamiento = Request.Form("Orden")

            Codigo = SecuenciaCodigo()
            Nombre = Request.Form("Nombre")
            Musica = Request.Form("Musica")
            Video = Request.Form("Video")
            Juegos = Request.Form("Juegos")
            Software = Request.Form("Software")
            Libros = Request.Form("Libros")
            Hardware = Request.Form("Hardware")

            sqlString = "INSERT INTO discos_Tipos(Usuario, Codigo, Nombre, Musica, Video, Juegos, Software, Libros, Hardware) " & _
                        "VALUES ('" & Usuario & "', '" & Codigo & "', '" & Nombre & "', " & Musica & ", " & Video & ", " & Juegos & ", " & Software & ", " & Libros & ", " & Hardware & ");"

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")
                cc.execute(sqlString)
            cc.close: set cc = nothing

            response.redirect "lista.asp?v=" & Ver & "&o=" & Orden
        %>    
    </body>
</html>