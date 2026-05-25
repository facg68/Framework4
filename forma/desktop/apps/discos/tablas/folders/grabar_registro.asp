<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            sub ResetearDefaults()
                dim cc, sqlString

                sqlString = "UPDATE discos_Carpetas " & _
                            "SET PorDefecto = 0 " & _
                            "WHERE Usuario = '" & Request.Cookies("Usuario") & "';"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                    cc.execute(sqlString)
                cc.close: set cc = nothing
            end sub
        %>
    </head>

    <body>
        <%
            dim con, t, tt, sqlString
            dim Usuario, Codigo, Nombre, Descripcion, PorDefecto

            Usuario = Request.Form("usuario")
            Codigo = Request.Form("Codigo")
            Nombre = Request.Form("Nombre")
            Descripcion = Request.Form("Descripcion")
            PorDefecto = Request.Form("PorDefecto")

            if PorDefecto = 1 then ResetearDefaults()

            sqlString = "UPDATE discos_Carpetas " & _
                        "SET Nombre = '" & Nombre & "', " & _
                           " Descripcion = '" & Descripcion & "', " & _
                           " PorDefecto = " & PorDefecto & " " & _
                     " WHERE (Usuario = '" & Usuario & "') " & _
                        "AND (Codigo = '" & Codigo & "');"

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")    
                con.execute(sqlString)
            con.close: set con = nothing

            response.redirect "lista.asp?o=" & Request.form("Orden")

        'response.write sqlString
        %>
    <body>
</html>