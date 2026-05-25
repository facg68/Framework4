<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim con, t, tt, sqlString
    dim Usuario, Codigo, Nombre, Estatus, Tipo, Ordenamiento, Juegos, Software, Obsoleta

    Usuario = Request.Cookies("usuario")
    Codigo = Request.Form("Codigo")
    Estatus = Request.Form("Estatus")
    Tipo = Request.Form("Tipo")
    Ordenamiento = Request.Form("Ordenamiento")

    Nombre = Request.Form("Nombre")
    Juegos = Request.Form("Juegos")
    Software = Request.Form("Software")
    Obsoleta = Request.Form("Obsoleta")

    sqlString = "UPDATE discos_Plataformas " & _
                "SET Nombre = '" & Nombre & "', " & _
                   " Juegos = " & Juegos & ", " & _
                   " Software = " & Software & ", " & _
                   " Obsoleta = " & Obsoleta & " " & _
            " WHERE (Usuario = '" & Usuario & "') " & _
                "AND (Codigo = '" & Codigo & "');"

    set con = Server.CreateObject("ADODB.Connection")
    con.open Application("Conn")    
        con.execute(sqlString)
    con.close: set con = nothing

    response.redirect "lista.asp?e=" & Estatus & "&t=" & Tipo & "&o=" & Ordenamiento
'response.write sqlString
        %>
    <body>
</html>