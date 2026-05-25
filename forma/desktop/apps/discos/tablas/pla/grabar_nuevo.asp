<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function SecuenciaCodigo()
                dim cc, tt, sqlString

                sqlString = "SELECT ISNULL(MAX(CAST(Codigo AS numeric(10, 0))), 0) + 1 AS Valor " & _
                            "FROM dbo.discos_Plataformas WHERE (Usuario = '" & Request.Cookies("Usuario") & "')"

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
            dim cc, tt, sqlString
            dim Usuario, Estatus, Tipo, Ordenamiento, Codigo, Nombre, Juegos, Software, Obsoleta

            Usuario = Request.Form("Usuario")
            Estatus = Request.Form("Estatus")
            Tipo = Request.Form("Tipo")
            Ordenamiento = Request.Form("Ordenamiento")

            Codigo = SecuenciaCodigo()
            Nombre = Request.Form("Nombre")
            Juegos = Request.Form("Juegos")
            Software = Request.Form("Software")

            sqlString = "INSERT INTO discos_Plataformas(Usuario, Codigo, Nombre, Juegos, Software, Obsoleta) " & _
                        "VALUES ('" & Usuario & "', '" & Codigo & "', '" & Nombre & "', " & Juegos & ", " & Software & ", 0);"

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")
                cc.execute(sqlString)
            cc.close: set cc = nothing

            response.redirect "lista.asp?e=" & Estatus & "&t=" & Tipo & "&o=" & Ordenamiento
        %>    
    </body>
</html>