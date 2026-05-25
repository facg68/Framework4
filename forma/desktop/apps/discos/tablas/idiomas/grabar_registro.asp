<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <body>
        <%
            dim con, t, tt, sqlString
            dim Usuario, Codigo, Nombre

            Usuario = Request.Cookies("usuario")
            Codigo = Request.Form("Codigo")
            Nombre = Request.Form("Nombre")

            sqlString = "UPDATE discos_Idiomas " & _
                           "SET Nombre = '" & Nombre & "' " & _
                         "WHERE (Usuario = '" & Usuario & "') " & _
                           "AND (Codigo = '" & Codigo & "');"

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")
                cc.execute(sqlString)
            cc.close: set cc = nothing

            response.redirect "lista.asp"
        %>    
    </body>
</html>