<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 

<!DOCTYPE html>

<html>
    <head>
        <%
            dim con, t, sqlString, nuevo
            dim codigo, tipocuenta, nombre, tipo, anualidad
            dim contacto, clase, categoria, grupo, localmonetario
            dim ValorMonto, mensajedefault, repetitiva

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            function unFormatNumber(Valor)
                dim v 

                v = trim(Valor)
                v = Replace(v, ",", "")   
                If Not IsNumeric(v) Then v = 0 

                unFormatNumber = v           
            end function     

            sub GrabarCuenta()
                if Nuevo = 1 then
                    if tipocuenta = "N" then
                        sqlString = "INSERT INTO pre_Cuentas(Usuario, Codigo, Nombre, Categoria, Tipo, Anualidad, Monto, Contacto," & _
                                                            " LocalMonetario, MensajeDefault, TipoCuenta, Repetitiva, DeSistema, Grupo, Clase) " & _
                                    "VALUES('" & lcase(Request.Cookies("Usuario")) & "', '" & codigo & "', '" & nombre & "', " & _
                                            "'" & categoria & "', '" & tipo & "', " & anualidad & ", '" & ValorMonto & "', '" & contacto & "', " & _
                                            "'" & localmonetario & "', '" & mensajedefault & "', '" & tipocuenta & "', " & repetitiva & ", 0, " & _
                                            "'" & grupo & "','" & clase & "');"
                    else
                        '
                        ' El Tipo "A" (Acumulador) no tiene los campos Anualidad, Clase, Monto
                        '
                        sqlString = "INSERT INTO pre_Cuentas(Usuario, Codigo, Nombre, Categoria, Tipo, Anualidad, Monto, Contacto," & _
                                                            " LocalMonetario, MensajeDefault, TipoCuenta, Repetitiva, DeSistema, Grupo, Clase) " & _
                                    "VALUES('" & lcase(Request.Cookies("Usuario")) & "', '" & codigo & "', '" & nombre & "', " & _
                                            "'" & categoria & "', '" & tipo & "', NULL, 0.00, '" & contacto & "', " & _
                                            "'" & localmonetario & "', '" & mensajedefault & "', '" & tipocuenta & "', " & repetitiva & ", 0, " & _
                                            "'" & grupo & "','N');"            
                    end if
                else
                    if tipocuenta = "N" then
                        sqlString = "UPDATE pre_Cuentas " & _
                                        "SET Nombre = '" & nombre & "', " & _
                                            " Categoria = '" & categoria & "', " & _
                                            " Tipo = '" & tipo & "', " & _ 
                                            " Anualidad = " & anualidad & ", " & _
                                            " Monto = '" & ValorMonto & "', " & _
                                            " Contacto = '" &  contacto & "', " & _
                                            " LocalMonetario = '" & localmonetario & "', " & _
                                            " MensajeDefault = '" & mensajedefault & "', " & _
                                            " TipoCuenta = '" & tipocuenta & "', " & _
                                            " Repetitiva = " & repetitiva & ", " & _
                                            " Grupo = '" & grupo & "', " & _
                                            " Clase = '" & clase & "' " & _
                                        "WHERE (Usuario = '" & request.Cookies("usuario") & "') " & _
                                        "AND (Codigo = '" & codigo & "');"
                    else
                        sqlString = "UPDATE pre_Cuentas " & _
                                        "SET Nombre = '" & nombre & "', " & _
                                            " Categoria = '" & categoria & "', " & _
                                            " Tipo = '" & tipo & "', " & _
                                            " Contacto = '" &  contacto & "', " & _
                                            " LocalMonetario = '" & localmonetario & "', " & _
                                            " MensajeDefault = '" & mensajedefault & "', " & _
                                            " TipoCuenta = '" & tipocuenta & "', " & _
                                            " Repetitiva = " & repetitiva & ", " & _
                                            " Grupo = '" & grupo & "' " & _
                                        "WHERE (Usuario = '" & request.Cookies("usuario") & "') " & _
                                        "AND (Codigo = '" & codigo & "');"
                    end if
                end if

                con.execute(sqlString)  
            end sub             
        %>           
    </head>

    <body>
        <%
            nuevo = Request.Form("Nuevo")

            if nuevo = 1 then
                codigo = Request.Form("codigo")
            else
                codigo = Request.Form("cod")
            end if

            tipocuenta = Request.Form("tipocuenta")
            nombre = Request.Form("nombre")
            tipo = Request.Form("tipo")
            anualidad = Request.Form("anualidad")
            contacto = Request.Form("contacto")
            clase = Request.Form("clase")
            categoria = Request.Form("categoria")
            grupo = Request.Form("grupo")
            localmonetario = Request.Form("localmonetario")
            ValorMonto = unFormatNumber(Request.Form("Monto"))
            mensajedefault = Request.Form("mensajedefault")
            repetitiva = Request.Form("repetitiva")     

            if isnull(Request.Form("anualidad")) or (Request.Form("anualidad") = "") then
                anualidad = "NULL"
            else
                anualidad = "'" & Request.Form("anualidad") &  "'"
            end if

            GrabarCuenta

            con.close: set con = nothing

            response.redirect "lista.asp"
        %>
    </body>
</html>