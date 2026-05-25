<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function SecuenciaCodigo()
                dim cc, tt, sqlString

                sqlString = "SELECT ISNULL(MAX(CAST(Forma AS numeric(10, 0))), 0) + 1 AS Valor " & _
                            "FROM dbo.discos_Formas WHERE (Usuario = '" & Request.Cookies("Usuario") & "')"

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
            dim con, t, tt, sqlString, grupo
            dim Usuario, Forma, Nombre, Multilados, Musica, Video, Juegos, Software, Libros, Hardware, Estatus

            Usuario = Request.Cookies("usuario")
            Nombre = Request.Form("Nombre")
            Multilados = Request.Form("Multilados")
            Musica = Request.Form("Musica")
            Video = Request.Form("Video")
            Juegos = Request.Form("Juegos")
            Software = Request.Form("Software")
            Libros = Request.Form("Libros")
            Hardware = Request.Form("Hardware")
            Estatus = Request.Form("Estatus")

            sqlString = "INSERT INTO discos_Formas(Usuario, Forma, Nombre, Multilados, Musica, Video, Juegos, Software, Libros, Hardware, Estatus) " & _
                        "VALUES('" & Usuario & "', '" & SecuenciaCodigo() & "', '" & Nombre & "', " & Multilados & ", " & Musica & ", " & Video & ", " & Juegos & _
                              ", " & Software & ", " & Libros & ", " & Hardware & ", " & Estatus & ");"

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")    
                con.execute(sqlString)
            con.close: set con = nothing

            response.redirect "lista.asp?g=" & request.form("grupo")
        %>    
    </body>
</html>