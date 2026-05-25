<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim con, t, tt, sqlString, grupo
    dim Usuario, Forma, Nombre, Multilados, Musica, Video, Juegos, Software, Libros, Hardware, Estatus

    Usuario = Request.Cookies("usuario")
    Forma = Request.Form("Forma")
    Nombre = Request.Form("Nombre")
    Multilados = Request.Form("Multilados")
    Musica = Request.Form("Musica")
    Video = Request.Form("Video")
    Juegos = Request.Form("Juegos")
    Software = Request.Form("Software")
    Libros = Request.Form("Libros")
    Hardware = Request.Form("Hardware")
    Estatus = Request.Form("Estatus")

    sqlString = "UPDATE discos_Formas " & _
                "SET Nombre = '" & Nombre & "', " & _
                " Multilados = " & Multilados & ", " & _
                " Musica = " & Musica & ", " & _
                " Video = " & Video & ", " & _
                " Juegos = " & Juegos & ", " & _
                " Software = " & Software & ", " & _
                " Libros = " & Libros & ", " & _
                " Hardware = " & Hardware & ", " & _
                " Estatus = " & Estatus & _
            " WHERE (Usuario = '" & Usuario & "') " & _
                "AND (Forma = '" & Forma & "');"

    set con = Server.CreateObject("ADODB.Connection")
    con.open Application("Conn")    
        con.execute(sqlString)
    con.close: set con = nothing

    response.redirect "lista.asp?g=" & Request.Form("Grupo")
%>