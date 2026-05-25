<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function FechaServer(FechaForm)
                dim d, m, a

                if FechaForm = "" then 
                    FechaServer = "NULL"
                else
                    d = left(FechaForm, 2)
                    m = mid(FechaForm, 4, 2)
                    a = right(FechaForm, 4)

                    FechaServer = a & "-" & right("00" & m, 2) & "-" & right("00" & d, 2)
                    FechaServer = "'" & FechaServer & "'"
                end if
            end function           
        %>
    </head>

    <body>
        <%
            dim c, t, sqlString, Cuenta, MultiPrecio

            set c = Server.CreateObject("ADODB.Connection")
            c.open Application("Conn")

                Cuenta = Request.Form("Cuenta")
                MultiPrecio = Request.form("MultiPrecio")

                sqlString = "SELECT Secuencia, RIGHT('00000000000000000000' + CAST(d.Secuencia AS varchar(18)), 18) AS Llave " & _
                                "FROM pre_Listas_Detalles AS d " & _
                                "WHERE (d.Usuario = '" & Request.Form("Usuario") & "') " & _
                                "AND (d.Codigo = '" & Request.Form("Codigo") & "') " & _
                            "ORDER BY d.Item;"

                set t = c.execute(sqlString)
                    if not (t.bof or t.eof) then
                        cuantos = 0   
                        sqlDetalle = ""

                        do
                            nombreItem           = "Litem_" & t("Llave")
                            nombrePrecioOriginal = "Lpori_" & t("Llave")
                            nombrePrecio         = "Lprec_" & t("Llave")
                            nombreFecha          = "Lfech_" & t("Llave")

                            if Cuenta = 1 then 
                                If Multiprecio = 0 then
                                    TipoDetalle = 1
                                else
                                    TipoDetalle = 2
                                end if                
                            else
                                If Multiprecio = 0 then
                                    TipoDetalle = 3
                                else
                                    TipoDetalle = 4
                                end if
                            end if                               

                            Select Case TipoDetalle
                                Case 1
                                    valorItem = request.form(nombreItem)
                                    valorPrecioOriginal = request.form(nombrePrecioOriginal)

                                    sqlDetalle = "UPDATE pre_Listas_Detalles " & _
                                                    "SET Item = '" & valorItem & "', " & _
                                                       " PrecioOriginal = " & valorPrecioOriginal & _
                                               " WHERE Secuencia = " & t("Secuencia") & ";"

                                Case 2
                                    valorItem = request.form(nombreItem)
                                    valorPrecioOriginal = request.form(nombrePrecioOriginal)                                              
                                    valorPrecio = request.form(nombrePrecio)     

                                    sqlDetalle = "UPDATE pre_Listas_Detalles " & _
                                                    "SET Item = '" & valorItem & "', " & _
                                                       " PrecioOriginal = " & valorPrecioOriginal & ", " & _
                                                       " Precio = " & valorPrecio & _
                                               " WHERE Secuencia = " & t("Secuencia") & ";" 

                                Case 3
                                    valorItem = request.form(nombreItem)
                                    valorPrecioOriginal = request.form(nombrePrecioOriginal)
                                    valorFecha = FechaServer(request.form(nombreFecha))

                                    sqlDetalle = "UPDATE pre_Listas_Detalles " & _
                                                    "SET Item = '" & valorItem & "', " & _
                                                       " PrecioOriginal = " & valorPrecioOriginal & ", " & _
                                                       " Fecha = " & valorFecha & _
                                               " WHERE Secuencia = " & t("Secuencia") & ";"                                                                                                                 

                                Case 4
                                    valorItem = request.form(nombreItem)
                                    valorPrecioOriginal = request.form(nombrePrecioOriginal)
                                    valorPrecio = request.form(nombrePrecio)
                                    valorFecha = FechaServer(request.form(nombreFecha))

                                    sqlDetalle = "UPDATE pre_Listas_Detalles " & _
                                                    "SET Item = '" & valorItem & "', " & _
                                                       " PrecioOriginal = " & valorPrecioOriginal & ", " & _
                                                       " Precio = " & valorPrecio & ", " & _
                                                       " Fecha = " & valorFecha & _
                                               " WHERE Secuencia = " & t("Secuencia") & ";"                                                                                                                 
                            End Select                    

                            c.execute(sqlDetalle)
                            t.MoveNext
                        loop until t.eof
                    end if       
                    
                t.close: set t = nothing        
            c.close: set c = nothing

            response.redirect "lista.asp"
        %>    
    </body>
</html>


