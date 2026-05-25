<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            dim con, sqlString, p, NombreImagenNueva

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            function fechaServer(fechaFormulario)
                dim d, m, a, h, min

                d = right("00" & left(fechaFormulario, 2), 2)
                m = right("00" & mid(fechaFormulario, 4, 2), 2)
                a = mid(fechaFormulario, 7, 4)

                h = right("00" & mid(fechaFormulario, 12, 2), 2)
                min = right("00" & right(fechaFormulario, 2), 2)

                fechaServer = a & "-" & m & "-" & d & " " & h & ":" & min
            end function

            function NuevoCodigo()
                dim tt, maximo, sqlCommand

                sqlCommand = "SELECT ( MAX( CAST( RIGHT(Codigo, 9) AS Numeric(9, 0) ) ) + 1 ) AS M " & _
                               "FROM dbo.seg_Anuncios;"

                set t = con.execute(sqlCommand)
                    if (t.bof or t.eof) then
                        NuevoCodigo = "a000000001"
                    else
                        NuevoCodigo = "a" & RIGHT("0000000000" & t("M"), 9)
                    end if
                t.close: set t = nothing
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

            function SecAnuncio(CodigoA)                         
                dim tt, sqlCommand

                sqlCommand = "SELECT Secuencia " & _
                               "FROM dbo.seg_Anuncios " & _
                              "WHERE (Propietario = '" & Request.Cookies("Usuario") & "') " & _
                                "AND (Codigo = '" & CodigoA & "');"

                set t = con.execute(sqlCommand)
                    SecAnuncio = t("Secuencia")
                t.close: set t = nothing            
            end function  

            sub CopiarImagen(ImagenOriginal, NuevaImagen)
                dim FSO, miPath, Original, Copia, Nombre, extension

                Set FSO = Server.CreateObject("Scripting.FileSystemObject")
                miPath = Server.MapPath("/") & "\imagenes\anuncios\"

                '
                ' Extraemos la extensión
                '

                tam = len(ImagenOriginal)
                donde = -1

                for k = tam to 1 step -1
                    if mid(ImagenOriginal, k, 1) = "." then
                        if donde = -1 then
                            donde = k
                        end if
                    end if
                next       

                nombre = left(ImagenOriginal, (donde - 1))
                extension = right(ImagenOriginal, (tam - donde))  
                NombreImagenNueva = NuevaImagen & "." & extension

                Original = miPath & ImagenOriginal
                Copia = miPath & NombreImagenNueva

                FSO.CopyFile Original, Copia
            end sub  
        %>
    </head>

    <body>
        <%
            '
            ' 01. Recibimos los datos...
            '

            dim Usuario, titulo, desde, hasta, tipo, verPantalla, estatusAnuncio, ordenadoPor, NombreImagen, CodigoInterno
            dim escalarImagen, codigoHTML, segundos, CodigoAnuncio, sAnuncio, Secuencia, republicar, origen

            Usuario = Request.Cookies("Usuario")
            titulo = NullValue(Request.Form("Titulo"))
            desde = fechaServer(Request.Form("inicio"))
            hasta = fechaServer(Request.Form("fin"))
            tipo = Request.Form("tipo")                     
            escalarImagen = Request.Form("EscalarImagen")            
            codigoHTML = NullValue(LimpiarApostrofes(Request.form("Cuerpo")))
            segundos = Request.Form("Segundos")

            Secuencia = Request.Form("Secuencia")
            CodigoAnuncio = NuevoCodigo()
            republicar = Request.Form("republicar")
            origen = Request.Form("origen")

            verPantalla = Request.Form("verPantalla")
            estatusAnuncio = Request.Form("estatusAnuncio")
            ordenadoPor = Request.Form("ordenadoPor")           

            '
            ' 02. Grabar Registro
            '
            if secuencia = "*" then
                '
                ' Nueva Publicación
                '
                sqlstring = "INSERT INTO seg_anuncios(Propietario, Titulo, Inicio, Fin, Tipo, EscalarImagen, Segundos, Cuerpo, Codigo) " & _
                            "VALUES('" & Usuario & "', " & titulo & ", '" & desde & "', '" & hasta & "', " & tipo & ", " & EscalarImagen & ", " & Segundos & ", " & codigoHTML & ", '" & CodigoAnuncio & "');"
            else
                if republicar = 0 then
                    '
                    ' Actualizamos una Publicación
                    '
                    sqlString = "UPDATE seg_anuncios " & _
                                "SET Titulo = " & titulo & "," & _
                                    " Inicio = '" & desde & "'," & _
                                    " Fin = '" & hasta & "'," & _
                                    " Tipo = " & tipo & "," & _
                                    " EscalarImagen = " & EscalarImagen & "," & _
                                    " Segundos = " & Segundos & "," & _
                                    " Cuerpo = " & codigoHTML & " " & _
                                "WHERE (Secuencia = " & secuencia & ");"                 
                else
                    '
                    ' Creamos un nuevo registro, pero conuna copia de la imagen
                    '
                    if len(trim(Request.Form("NombreImagen"))) > 0 then
                        CopiarImagen Request.Form("NombreImagen"), CodigoAnuncio
                    end if                        

                    sqlstring = "INSERT INTO seg_anuncios(Propietario, Titulo, Inicio, Fin, Tipo, EscalarImagen, Segundos, Cuerpo, Codigo, NombreImagen) " & _
                                "VALUES('" & Usuario & "', " & titulo & ", '" & desde & "', '" & hasta & "', " & tipo & ", " & EscalarImagen & ", " & Segundos & ", " & codigoHTML & ", '" & CodigoAnuncio & "', '" & NombreImagenNueva & "');"

                end if
            end if

            con.execute(sqlString)


            '
            ' 03. Grabar las Pantallas
            '

            if (secuencia = "*") OR (republicar = 1) then
                sAnuncio = SecAnuncio(CodigoAnuncio)
            else
                con.execute("DELETE FROM seg_Anuncios_Asignaciones WHERE Anuncio = " & secuencia & ";")            
                sAnuncio = Secuencia
            end if

            set p = con.execute("SELECT Pantalla, Nombre FROM seg_Anuncios_Pantallas ORDER BY Nombre;")

            if not (p.bof or p.eof) then
                Do
                    nombreCampo = "pantalla_" & Trim(p("Pantalla"))
                    valor = Request.form(nombreCampo)

                    if valor = 1 then
                        '
                        ' Insertamos el registro en la lista de pantallas...
                        '
                        sqlString = "INSERT INTO seg_Anuncios_Asignaciones(Anuncio, Pantalla) " & _
                                    "VALUES (" & sAnuncio & ", '" & p("Pantalla") & "');"

                        con.execute(sqlString)
                    end If 

                    p.MoveNext
                Loop Until (p.eof)
            end if

            p.close: set p = nothing
            con.close: set con = nothing


            '
            ' 04. Subir el archivo (si no es HTML)
            '
            ' Creamos un FORMULARIO para RE-ENVIAR 
            ' la informacion del archivo seleccionado
            '

            if (Tipo = 1) or (Tipo = 3)  then
                if secuencia = "*" then
                    '
                    ' Solo se hace automáticamente al crear nueva publicación
                    '
                    response.redirect "subir_objeto.asp?w=" & origen & "&a=" & sAnuncio & "&tv=" & verPantalla & "&e=" & estatusAnuncio & "&op=" & ordenadoPor
                else
                    if origen = 1 then
                        response.redirect "lista_total.asp?tv=" & verPantalla & "&e=" & estatusAnuncio & "&op=" & ordenadoPor
                    else
                        response.redirect "lista.asp?tv=" & verPantalla & "&e=" & estatusAnuncio & "&op=" & ordenadoPor
                    end if
                end if
            else
                    if origen = 1 then
                        response.redirect "lista_total.asp?tv=" & verPantalla & "&e=" & estatusAnuncio & "&op=" & ordenadoPor
                    else
                        response.redirect "lista.asp?tv=" & verPantalla & "&e=" & estatusAnuncio & "&op=" & ordenadoPor
                    end if            
            end if      
        %>    
    </body>
</html>