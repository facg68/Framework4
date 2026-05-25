<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function LocalUsuario()
                dim cc, tt, sqlString

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.execute("SELECT usuLocal FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';")

                    if tt("usuLocal") = "" then
                        LocalUsuario = "US"
                    else
                        LocalUsuario = tt("usuLocal")
                    end if

                tt.close: set tt = nothing
                cc.close: set c = nothing
            end function
        %>
    </head>

    <body>
        <%
            dim c, sqlString
            dim Usuario, Codigo, Nombre, Descripcion

            Usuario = Request.Cookies("Usuario")
            Codigo = Request.Form("nuevoCodigo")
            Nombre = Request.Form("nuevoNombre")
            Descripcion = Request.Form("nuevaDescripcion")

            sqlString = "INSERT INTO pre_Listas_Encabezado(Usuario, Codigo, Nombre, Descripcion, Cuenta, Monto, Contacto," & _
                                " PrecioOriginal, PrecioFinal, MultiPrecio, Grupo, Categoria, VerListaEnInforme) " & _
                        "VALUES('" & Usuario & "', '" & Codigo & "', '" & Nombre & "', '" & Descripcion & "', 0, 0.00, NULL," & _
                                "'" &  LocalUsuario() & "', '" &  LocalUsuario() & "', 0, 'A', 'compras', 0);"

            set c = Server.CreateObject("ADODB.Connection")
            c.open Application("Conn")
                c.execute(sqlString)
            c.close: set c = nothing

            response.redirect "lista.asp"
        %>
    </body>
</html>