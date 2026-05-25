<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            dim con, sqlString

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")
            
            function fechaServer(fechaFormulario)
                dim d, m, a, h, min

                d = right("00" & left(fechaFormulario, 2), 2)
                m = right("00" & mid(fechaFormulario, 4, 2), 2)
                a = mid(fechaFormulario, 7, 4)

                h = right("00" & mid(fechaFormulario, 12, 2), 2)
                min = right("00" & right(fechaFormulario, 2), 2)

                fechaServer = "'" & a & "-" & m & "-" & d & " " & h & ":" & min & "'"
            end function

            function LimpiarApostrofes(valor)
                LimpiarApostrofes = Replace(valor,"'","´")
            end function 

            function NullValue(Cadena)
                if isNull(Cadena) then
                    NullValue = "NULL"
                else 
                    if len(trim(Cadena)) = 0 then
                        NullValue = "NULL"
                    else
                        NullValue = "'" & Cadena & "'"
                    end if   
                end if
            end function   

            Function CodigoUnico()
                Dim ahora, codigo, aleatorio
                Randomize
                ahora = Now
                aleatorio = Right("000" & Int(Rnd * 1000), 3) ' 3 dígitos pseudo-milisegundos
                codigo = Year(ahora) & _
                        Right("0" & Month(ahora), 2) & _
                        Right("0" & Day(ahora), 2) & _
                        Right("0" & Hour(ahora), 2) & _
                        Right("0" & Minute(ahora), 2) & _
                        Right("0" & Second(ahora), 2) & _
                        aleatorio
                CodigoUnico = codigo
            End Function           
        %>
    </head>

    <body>
        <%
            '
            ' 01. Recibimos los datos...
            '

            dim Secuencia, Nombre, Desde, Hasta, CU

            Usuario = Request.Cookies("Usuario")

            Secuencia = Request.Form("Secuencia")
            Nombre = NullValue(Request.Form("Nombre"))
            Desde = fechaServer(Request.Form("Desde"))
            Hasta = fechaServer(Request.Form("Hasta"))

            estatusAnuncio = Request.Form("estatusAnuncio")
            ordenadoPor = Request.Form("ordenadoPor")           

            CU = CodigoUnico()

            '
            ' 02. Grabar Registro
            '
            if secuencia = "*" then
                '
                ' Nuevo Panfleto
                '
                sqlstring = "INSERT INTO seg_Panfletos(CreadoPor, Nombre, PublicarDesde, PublicarHasta, CU) " & _
                            "VALUES('" & Usuario & "', " & Nombre & ", " & Desde & ", " & Hasta & ", '" & CU & "');"
            else
                '
                ' Actualizamos Panfleto
                '
                sqlString = "UPDATE seg_Panfletos " & _
                            "SET CreadoPor = '" & Usuario & "'," & _
                                " Nombre = " & Nombre & "," & _
                                " PublicarDesde = " & Desde & "," & _
                                " PublicarHasta = " & Hasta & _
                            " WHERE (Secuencia = " & secuencia & ");"                 
            end if

            response.write sqlString & "<br />"
            con.execute(sqlString)

            '
            ' 04. Subir el archivo (si no es HTML)
            '
            ' Creamos un FORMULARIO para RE-ENVIAR 
            ' la informacion del archivo seleccionado
            '

            if secuencia = "*" then
                '
                ' Solo se hace automáticamente al crear nueva Panfleto
                '
                response.redirect "subir_objeto.asp?cu=" & CU & "&e=" & estatusAnuncio & "&op=" & ordenadoPor
            else
                response.redirect "lista.asp?e=" & estatusAnuncio & "&op=" & ordenadoPor
            end if
        %>    
    </body>
</html>