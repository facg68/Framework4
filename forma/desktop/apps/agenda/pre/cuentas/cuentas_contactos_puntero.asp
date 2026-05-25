<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function sqlNow()
                dim p1, p2, p3, k, p, cadena, segmento, inicio

                '
                ' Parsear a mano... Uff!!!
                '

                p = 0
                cadena = FormatDateTime(Now(), 2) & "/"
                inicio = 1

                for k = 1 to len(cadena)
                if mid(cadena, k, 1) = "/" then
                    p = p + 1
                    segmento = mid(cadena, inicio, (k - inicio))
                    inicio = k + 1

                    select case p
                    case 1: p1 = right("0" & segmento, 2)
                    case 2: p2 = right("0" & segmento, 2)
                    case 3: p3 = segmento
                    end select
                end if
                next

                sqlNow = p3 & "-" & p1 & "-" & p2 & " " & FormatDateTime(Now(), 4)
            end function              
        %>
    </head>

    <body>
        <%
            dim c, t, sqlString, Secuencia
            dim Usuario, Cuenta, Contacto, Monto

            Secuencia = Request.QueryString("s")

            set c = Server.CreateObject("ADODB.Connection")
            c.open Application("Conn")
            set t = c.execute("SELECT * FROM pre_Cuentas_Comparticiones WHERE Secuencia = " & Secuencia & ";")

                Usuario = Request.Cookies("Usuario")
                Cuenta = t("Cuenta")

                '
                ' Actualizamos las fecha-hora de todos los usuarios a NULL
                '
                sqlString = "UPDATE pre_Cuentas_Comparticiones " & _
                            "SET UltimaFechaAplicada = NULL," & _
                               " Puntero = 0 " & _
                            " WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Cuenta = '" & Cuenta & "');"

                c.execute(sqlString)

                '
                ' Actualizamos la fecha-hora del contacto indicado...
                '
                sqlString = "UPDATE pre_Cuentas_Comparticiones " & _
                            "SET UltimaFechaAplicada = '" & sqlNow() & "'," & _
                               " Puntero = 1 " & _
                            " WHERE Secuencia = " & Secuencia & ";"

                c.execute(sqlString)

            '
            ' Cerramos la conexión y volvemos al editor...
            '
            c.close: set c = nothing

            response.redirect "cuentas_contactos.asp?c=" & cuenta
        %>
    </body>
</html>