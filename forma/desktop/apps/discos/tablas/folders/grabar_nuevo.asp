<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function SecuenciaCodigo()
                dim cc, tt, sqlString

                sqlString = "SELECT ISNULL(MAX(CAST(Codigo AS numeric(10, 0))), 0) + 1 AS Valor " & _
                            "FROM dbo.discos_Carpetas WHERE (Usuario = '" & Request.Cookies("Usuario") & "')"

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
            dim Usuario, Codigo, Nombre, Descripcion, PorDefecto

            Usuario = Request.Form("usuario")
            Codigo = SecuenciaCodigo()
            Nombre = Request.Form("Nombre")
            Descripcion = Request.Form("Descripcion")
            PorDefecto = Request.Form("PorDefecto")

            sqlString = "INSERT INTO discos_Carpetas(Usuario, Codigo, Nombre, Descripcion, PorDefecto) " & _
                        "VALUES ('" & Usuario & "', '" & Codigo & "', '" & Nombre & "', '" & Descripcion & "', " &  PorDefecto & ");"

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")    
                con.execute(sqlString)
            con.close: set con = nothing

            response.redirect "lista.asp?o=" & Request.form("Orden")
        %>
    <body>