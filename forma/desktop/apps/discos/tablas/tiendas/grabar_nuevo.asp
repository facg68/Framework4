<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function SecuenciaCodigo()
                dim cc, tt, sqlString

                sqlString = "SELECT ISNULL(MAX(CAST(Codigo AS numeric(10, 0))), 0) + 1 AS Valor " & _
                            "FROM dbo.discos_Tiendas WHERE (Usuario = '" & Request.Cookies("Usuario") & "')"

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
            dim cc, tt, sqlString, Grupo
            dim Usuario, Codigo, Nombre, Contacto, SitioWeb, Correo, Tipo, Pais
            dim Telefono1, Telefono2, Direccion, Notas, MediosDigitales, MediosFisicos
            dim Musica, Video, Juegos, Software, Libros, Hardware, Estatus

            Usuario = Request.Cookies("Usuario")
            Codigo = SecuenciaCodigo()
            Grupo = Request.Form("Grupo")

            Nombre = Request.Form("Nombre")
            Contacto = Request.Form("Contacto")
            SitioWeb = Request.Form("SitioWeb")
            Correo =Request.Form("Correo")
            Tipo = Request.Form("Tipo")
            Pais = Request.Form("Pais")
            Telefono1 = Request.Form("Telefono1")
            Telefono2 = Request.Form("Telefono2")
            Direccion = Request.Form("Direccion")
            Notas = Request.Form("Notas")
            MediosDigitales = Request.Form("MediosDigitales")
            MediosFisicos = Request.Form("MediosFisicos")
            Musica = Request.Form("Musica")
            Video = Request.Form("Video")
            Juegos = Request.Form("Juegos")
            Software = Request.Form("Software")
            Libros = Request.Form("Libros")
            Hardware = Request.Form("Hardware")

            sqlString = "INSERT INTO discos_Tiendas (Usuario, Codigo, Nombre, Contacto, SitioWeb, Correo, Tipo, Pais, Telefono1, Telefono2, " & _
                        "Direccion, Notas, MediosDigitales, MediosFisicos, Musica, Video, Juegos, Software, Libros, Hardware, Estatus) " & _
                        "VALUES ('" & Usuario & "', '" & Codigo & "', '" & Nombre & "', '" & Contacto & "', '" & SitioWeb & "', '" & Correo & "'," & Tipo & ", '" & Pais & "', '" & Telefono1 & "', '" & Telefono2 & "', " & _
                                 "'" & Direccion & "', '" & Notas & "'," & MediosDigitales & "," & MediosFisicos & "," & Musica & "," & Video & "," & Juegos & "," & Software & "," & Libros & ", " & Hardware & ", 1);"

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")
                cc.execute(sqlString)
            cc.close: set cc = nothing

            response.redirect "lista.asp?g=" & Grupo
        %>    
    </body>
</html>

