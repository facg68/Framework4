<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function SecuenciaCodigo()
                dim cc, tt, sqlString

                sqlString = "SELECT ISNULL(MAX(CAST(Codigo AS numeric(10, 0))), 0) + 1 AS Valor " & _
                            "FROM dbo.discos_Idiomas WHERE (Usuario = '" & Request.Cookies("Usuario") & "')"

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
            dim con, t, tt, sqlString
            dim Usuario, Codigo, Nombre

            Usuario = Request.Cookies("usuario")
            Codigo = SecuenciaCodigo()
            Nombre = Request.Form("Nombre")

            sqlString = "INSERT INTO discos_Idiomas(Usuario, Codigo, Nombre) " & _
                        "VALUES ('" & Usuario & "', '" & Codigo & "', '" & Nombre & "');"

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")
                cc.execute(sqlString)
            cc.close: set cc = nothing

            response.redirect "lista.asp"
        %>    
    </body>
</html>