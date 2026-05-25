<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
    </head>

    <body>
        <%    
            dim cc, tt, sqlString, Grupo
            dim Usuario, Codigo, Nombre, Contacto, SitioWeb, Correo, Tipo, Pais
            dim Telefono1, Telefono2, Direccion, Notas, MediosDigitales, MediosFisicos
            dim Musica, Video, Juegos, Software, Libros, Hardware, Estatus

            Usuario = Request.Cookies("Usuario")
            Codigo = Request.Form("Codigo")
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
            Estatus = Request.Form("Estatus")

            sqlString = "UPDATE discos_Tiendas " & _
                        "SET Nombre = '" & Nombre & "', " & _
                           " Contacto = '" & Contacto & "', " & _
                           " SitioWeb = '" & SitioWeb & "', " & _
                           " Correo = '" & Correo & "', " & _
                           " Tipo = " & Tipo & ", " & _
                           " Pais = '" & Pais & "', " & _
                           " Telefono1 = '" & Telefono1 & "', " & _
                           " Telefono2 = '" & Telefono2 & "', " & _
                           " Direccion = '" & Direccion & "', " & _
                           " Notas = '" & Notas & "', " & _
                           " MediosDigitales = " & MediosDigitales & ", " & _
                           " MediosFisicos = " & MediosFisicos & ", " & _
                           " Musica = " & Musica & ", " & _
                           " Video = " & Video & ", " & _
                           " Juegos = " & Juegos & ", " & _
                           " Software = " & Software & ", " & _
                           " Libros = " & Libros & ", " & _
                           " Hardware = " & Hardware & ", " & _
                           " Estatus = " & Estatus & _
                    " WHERE (Usuario = '" & Usuario & "') " & _
                       "AND (Codigo = '" & Codigo & "');"

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")    
                con.execute(sqlString)
            con.close: set con = nothing

            response.redirect "lista.asp?g=" & Grupo

'response.write sqlString
        %>
    <body>
</html>