<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            Function SecuenciaPaquete()
                dim cc, tt, sqlString, secuencia

                sqlString = "select ISNULL(MAX(CAST(RIGHT(paquete, 6) AS numeric(6,0))), 0) AS Valor " & _
                            "from discos_Paquetes " & _
                            "where Usuario ='" & Request.Cookies("Usuario") & "';"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.execute(sqlString)

                if not (tt.bof or tt.eof) then
                    secuencia = cDbl(tt("Valor")) + 1
                    SecuenciaPaquete = "PK" & RIGHT("000000" & secuencia, 6)
                end if

                tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function
        %>
    </head>

    <body>
        <%
            dim con, sqlString, Usuario, Amo, Paquete

            Usuario = Request.Cookies("Usuario")
            Amo = Year(Now())
            Paquete = SecuenciaPaquete()

            sqlString = "INSERT INTO discos_Paquetes(Usuario, Paquete, ACompra, AEdicion, Titulo, Precio, Tienda, Casa, Carpeta) " & _
                        "VALUES('" & Usuario & "', '" & Paquete & "', " & Amo & ", " & Amo & ",'Nuevo Medio', 0.00, '00000000', '00000000', '00000000');"

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")
            con.execute(sqlString)
            con.close: set con = nothing

            response.redirect "editar.asp?m=" & Paquete
        %>
    </body>
</html>