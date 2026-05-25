<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            Function FechaServer(fecha)
                Dim partes, d, m, y

                If (Trim(fecha) = "") or (len(fecha) <> 10) Then
                    FechaServer = ""
                    Exit Function
                End If

                partes = Split(fecha, "/")

                d = Right("0" & partes(0), 2)
                m = Right("0" & partes(1), 2)
                y = partes(2)

                FechaServer = y & "-" & m & "-" & d
            End Function            

            function unFormatNumber(Valor)
                dim cadena, k, char

                cadena = cStr(Valor)
                unFormatNumber = ""

                if Cadena <> "" then
                    for k = 1 to len(cadena)
                        char = mid(cadena, k, 1)
                        if char <> "," then
                            unFormatNumber = unFormatNumber & char
                        end if
                    next
                end if
            end function

            function LocalMonetarioUsuario(Monto, MonedaOrigen, MonedaDestino)
                dim con, t, sqlString

                sqlString = "SELECT dbo.Cripto_CambiarMoneda(" & Monto & ", '" & MonedaOrigen & "', '" & MonedaDestino & "') AS Cambio;"

                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")
                set t = con.execute(sqlString)

                Convertir = t("Cambio")

                t.close: set t = nothing
                con.close: set con = nothing
            End Function               

            sub ActualizarEncabezado(Usuario, Codigo)
                dim cc, tt, sqlString, Total

                sqlString = "SELECT SUM(d.PrecioOriginal) AS Monto " & _
                              "FROM pre_Listas_Detalles AS d " & _
                        "INNER JOIN pre_Listas_Encabezado AS e " & _
                                "ON d.Usuario = e.Usuario " & _
                               "AND d.Codigo = e.Codigo " & _
                             "WHERE (e.Usuario = '" & Usuario & "') " & _
                               "AND (e.Codigo = '" & Codigo & "');"
                
                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.execute(sqlString)
                    Total = tt("Monto")
                tt.close: set tt = nothing

                sqlString = "UPDATE pre_Listas_Encabezado " & _
                               "SET Monto = " & Total & _
                            " WHERE (Usuario = '" & Usuario & "') " & _
                               "AND (Codigo = '" & Codigo & "');"
                cc.execute(sqlString)

                cc.close: set cc = nothing
            end sub
        %>
    </head>

    <body>
        <%
            dim con, t, sqlString, t1, MultiPrecio, Cuenta, Usuario,lOriginal, lDestino
            dim codigo, NuevoItem, PrecioOriginal, Precio, Fecha

            Codigo = request.QueryString("cod")
            Usuario = Request.Cookies("Usuario")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            '-------------------------------------------------
            ' Abrimos la Lista para cargar los parámetros... '
            '-------------------------------------------------

            set t1 = con.execute("SELECT MultiPrecio, Cuenta, PrecioOriginal, PrecioFinal " & _
                                  "FROM pre_Listas_Encabezado " & _
                                  "WHERE (Usuario = '" & Usuario & "') " & _
                                  "AND (Codigo = '" & Codigo & "');")

              Multiprecio = t1("MultiPrecio")
              Cuenta = t1("Cuenta")
              lOriginal = t1("PrecioOriginal")
              lDestino = t1("PrecioFinal")

            t1.close: set t1 = nothing

            '--------------------
            ' Cerramos la Lista '
            '--------------------

            if Cuenta = 1 then
                if MultiPrecio = 1 then 
                    NuevoItem = Request.QueryString("i")
                    PrecioOriginal = unFormatNumber(Request.QueryString("p1"))
                    Precio = unFormatNumber(Request.QueryString("p2"))

                    sqlString = "INSERT INTO pre_Listas_Detalles(Usuario, Codigo, Item, PrecioOriginal, Precio) " & _
                                "VALUES('" & Usuario & "', '" & Codigo & "', '" & NuevoItem & "', " & PrecioOriginal & ", " & Precio & ";"
                else
                    NuevoItem = Request.QueryString("i")
                    PrecioOriginal = Request.QueryString("p1")

                    sqlString = "INSERT INTO pre_Listas_Detalles(Usuario, Codigo, Item, PrecioOriginal, Precio) " & _
                                "VALUES('" & Usuario & "', '" & Codigo & "', '" & NuevoItem & "', " & PrecioOriginal & _ 
                                ", dbo.Cripto_CambiarMoneda(" & PrecioOriginal & ", '" & lOriginal & "', '" & lDestino & "'));"
                end if
            else
                if MultiPrecio = 1 then 
                    NuevoItem = Request.QueryString("i")
                    PrecioOriginal = unFormatNumber(Request.QueryString("p1"))
                    Precio = unFormatNumber(Request.QueryString("p2"))
                    Fecha = Request.QueryString("f")

                    if Fecha = "" then
                        sqlString = "INSERT INTO pre_Listas_Detalles(Usuario, Codigo, Item, PrecioOriginal, Precio) " & _
                                    "VALUES('" & Usuario & "', '" & Codigo & "', '" & NuevoItem & "', " & PrecioOriginal & ", " & Precio & ");"
                    else
                        sFecha = FechaServer(Fecha)

                        sqlString = "INSERT INTO pre_Listas_Detalles(Usuario, Codigo, Item, PrecioOriginal, Precio, Fecha) " & _
                                    "VALUES('" & Usuario & "', '" & Codigo & "', '" & NuevoItem & "', " & PrecioOriginal & ", " & Precio & ", '" & sFecha & "');"
                    end if  
                else
                    NuevoItem = Request.QueryString("i")
                    PrecioOriginal = unFormatNumber(Request.QueryString("p1"))
                    Fecha = Request.QueryString("f")

                    if Fecha = "" then
                        sqlString = "INSERT INTO pre_Listas_Detalles(Usuario, Codigo, Item, PrecioOriginal, Precio) " & _
                                    "VALUES('" & Usuario & "', '" & Codigo & "', '" & NuevoItem & "', " & PrecioOriginal & _ 
                                    ", dbo.Cripto_CambiarMoneda(" & PrecioOriginal & ", '" & lOriginal & "', '" & lDestino & "'));"
                    else
                        sFecha = FechaServer(Fecha)

                        sqlString = "INSERT INTO pre_Listas_Detalles(Usuario, Codigo, Item, PrecioOriginal, Precio, Fecha) " & _
                                    "VALUES('" & Usuario & "', '" & Codigo & "', '" & NuevoItem & "', " & PrecioOriginal & _ 
                                    ", dbo.Cripto_CambiarMoneda(" & PrecioOriginal & ", '" & lOriginal & "', '" & lDestino & "'), '" & sFecha & "');"
                    end if
                end if
            end if

            con.execute(sqlString)
            con.close: set con = nothing

            ActualizarEncabezado Usuario, Codigo

            response.redirect "listas_items.asp?l=" & Codigo
        %>
    </body>
</html>
