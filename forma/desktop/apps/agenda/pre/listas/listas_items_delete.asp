<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            dim conn, sqlString, Codigo, Secuencia

            set conn = Server.CreateObject("ADODB.Connection")
            conn.open Application("Conn")

            function CodigoLista(SecuenciaItem)
                dim tt

                set tt = conn.execute("SELECT Codigo FROM pre_Listas_Detalles WHERE Secuencia = " & SecuenciaItem & ";")
                    CodigoLista = tt("Codigo")
                tt.close: set tt = nothing
            end function
        %>
    </head>

    <body>
        <%
            Secuencia = Request.QueryString("s")
            Codigo = CodigoLista(Secuencia)

            response.write "Secuencia = " & Secuencia & "<br />"
            response.write "Codigo = " & Codigo & "<br />"

            sqlString = "DELETE FROM pre_Listas_Detalles WHERE Secuencia = " & Secuencia & ";"
            conn.execute(sqlString)

            response.redirect "listas_items.asp?l=" & Codigo   

            conn.close: set conn = nothing
        %>
    </body>
</html>