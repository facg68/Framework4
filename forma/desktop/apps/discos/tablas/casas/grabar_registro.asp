<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim con, t, tt, sqlString
    dim Usuario, Codigo, Nombre, Musica, Video, Juegos, Software, Libros, Hardware, Obsoleta

    Usuario = Request.Cookies("usuario")
    Codigo = Request.Form("Codigo")
    Nombre = Request.Form("Nombre")
    Musica = Request.Form("Musica")
    Video = Request.Form("Video")
    Juegos = Request.Form("Juegos")
    Software = Request.Form("Software")
    Libros = Request.Form("Libros")
    Hardware = Request.Form("Hardware")
    Obsoleta = Request.Form("Obsoleta")

    sqlString = "UPDATE discos_Casas " & _
                "SET Nombre = '" & Nombre & "', " & _
                   " Musica = " & Musica & ", " & _
                   " Video = " & Video & ", " & _
                   " Juegos = " & Juegos & ", " & _
                   " Software = " & Software & ", " & _
                   " Libros = " & Libros & ", " & _
                   " Hardware = " & Hardware & ", " & _
                   " Obsoleta = " & Obsoleta & " " & _
            " WHERE (Usuario = '" & Usuario & "') " & _
                "AND (Codigo = '" & Codigo & "');"

    set con = Server.CreateObject("ADODB.Connection")
    con.open Application("Conn")    
        con.execute(sqlString)
    con.close: set con = nothing

    response.redirect "lista.asp?o=" & Request.form("Orden")
        %>
    <body>
</html>