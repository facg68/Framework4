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

            function TipoComparticion(Cuenta)
                dim fcon, f

                set fcon = Server.CreateObject("ADODB.Connection")
                fcon.Open Application("Conn")
                set f = fcon.execute("SELECT Clase " & _
                                    "FROM pre_Cuentas " & _
                                    "WHERE Usuario = '" & Request.Cookies("Usuario") & "' " & _
                                        "AND Codigo = '" & Cuenta & "';")

                TipoComparticion = f("Clase")
                
                f.close: set f = nothing
                fcon.close: set fcon = nothing
            end function      

            function CantidadCompartida(Cuenta)
                dim fcon, f

                set fcon = Server.CreateObject("ADODB.Connection")
                fcon.Open Application("Conn")
                set f = fcon.execute("SELECT Monto " & _
                                    "FROM pre_Cuentas " & _
                                    "WHERE Usuario = '" & Request.Cookies("Usuario") & "' " & _
                                        "AND Codigo = '" & Cuenta & "';")

                CantidadCompartida = CDbl(f("Monto"))
                
                f.close: set f = nothing
                fcon.close: set fcon = nothing
            end function  

            function TotalUsuarios(Cuenta)
                dim fcon, f

                set fcon = Server.CreateObject("ADODB.Connection")
                fcon.Open Application("Conn")
                set f = fcon.execute("SELECT COUNT(*) AS Cuantos " & _
                                    "FROM pre_Cuentas_Comparticiones " & _
                                    "WHERE Usuario = '" & Request.Cookies("Usuario") & "' " & _
                                        "AND Cuenta = '" & Cuenta & "';")

                TotalUsuarios = f("Cuantos")
                
                f.close: set f = nothing
                fcon.close: set fcon = nothing
            end function        

            function MontoUsuario(Cuenta)     
                Tipo = TipoComparticion(Cuenta)
                Monto = CantidadCompartida(Cuenta) 
                Usuarios = TotalUsuarios(Cuenta) - 1

                Select Case Tipo
                    Case "R"
                        MontoUsuario = Monto
                    Case "C"
                        if Usuarios < 1 then
                            MontoUsuario = Monto
                        else
                            MontoUsuario = (Monto / Usuarios)
                        end if
                    Case Else
                        MontoUsuario = 0.00
                End Select
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
                Monto = MontoUsuario(Cuenta)

                '
                ' Eliminamos el contacto...
                '
                sqlString = "DELETE FROM  pre_Cuentas_Comparticiones WHERE Secuencia = " & Secuencia & ";"
                c.execute(sqlString)

                '
                ' Actualizamos las cantidades de los usuarios compartidos...
                '
                sqlString = "UPDATE pre_Cuentas_Comparticiones " & _
                            "SET MontoCompartido = " & Monto & _
                            " WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Cuenta = '" & Cuenta & "');"

                c.execute(sqlString)

                '
                ' Cerramos la conexión y volvemos al editor...
                '
            t.close: set t = nothing
            c.close: set c = nothing

            response.redirect "cuentas_contactos.asp?c=" & cuenta
        %>
    </body>
</html>