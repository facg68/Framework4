<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <!-- Upload libre -->
        <!-- #include virtual = "/core/includes/freeaspupload.asp" -->

        <%
            sub actualizarfoto(Anuncio, Nombre)
                dim con, sqlstring

                sqlstring = "UPDATE seg_anuncios SET NombreImagen = '" & Nombre & "' WHERE Secuencia = " & Anuncio & ";"

                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")
                    con.execute sqlstring
                con.close: set con = nothing
            end sub

            function NombreImagen(secuencia)
                dim c, t, sqlStr

                sqlStr = "SELECT NombreImagen FROM seg_Anuncios WHERE Secuencia = " & Secuencia

                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                set t = c.execute(sqlStr)

                    NombreImagen = trim(t("NombreImagen"))

                t.close: set t = nothing
                c.close: set c = nothing      
            end function

            sub BorrarImagen(secuencia)
                On Error resume next

                dim imagen, path, objeto

                Objeto = NombreImagen(Secuencia)

                if len(trim(Objeto)) > 0 then
                    path = Server.MapPath("/forma/desktop/apps/anuncios/publicaciones") & "\" 
                    imagen = path & Objeto

                    Set fso = CreateObject("Scripting.FileSystemObject")
                        if fso.FileExists(imagen) then fso.DeleteFile imagen
                    set fso = nothing      
                end if   

                On Error Goto 0      
            end sub   

            Function Codigo(Anuncio)
                dim cc, tt, sqlstring

                sqlstring = "SELECT Codigo FROM seg_anuncios WHERE Secuencia = " & Anuncio

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.execute(sqlstring)
                    Codigo = tt("Codigo")
                tt.close: set tt = nothing
                cc.close: set cc = nothing                
            end function  

            Function ObjetoExtension(fOriginal)
                dim pos
                pos = InStrRev(fOriginal, ".")

                if pos > 0 then
                    ObjetoExtension = LCase(Mid(fOriginal, pos + 1))
                else
                    ObjetoExtension = ""
                end if
            end function      
        %>
    </head>

    <body>
        <%
            Dim Upload, uploadsDir
            Dim fileKey, NombreOriginal
            Dim a, basePath, tempPath, finalPath, extension
            Dim fso, nombreNuevo

            a = Request.Cookies("edit_anuncio")

            '----------------------------------------
            ' Borrar imagen anterior
            '----------------------------------------
            BorrarImagen a

            '----------------------------------------
            ' Carpeta destino
            '----------------------------------------
            uploadsDir = Server.MapPath("/forma/desktop/apps/anuncios/publicaciones")

            '----------------------------------------
            ' Upload
            '----------------------------------------
            If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

                Set Upload = New FreeASPUpload
                Upload.Save(uploadsDir)

                If Upload.UploadedFiles.Count = 0 Then
                    Response.Write "No se subió archivo"
                    Response.End
                End If

                ' Obtener archivo
                For Each fileKey In Upload.UploadedFiles.Keys
                    NombreOriginal = Upload.UploadedFiles(fileKey).FileName
                Next

                basePath = uploadsDir
                tempPath = basePath & "\" & NombreOriginal

                extension = ObjetoExtension(NombreOriginal)
                nombreNuevo = Codigo(a) & "." & extension
                finalPath = basePath & "\" & nombreNuevo

                '----------------------------------------
                ' Renombrar
                '----------------------------------------
                Set fso = CreateObject("Scripting.FileSystemObject")

                If fso.FileExists(finalPath) Then fso.DeleteFile finalPath
                fso.MoveFile tempPath, finalPath

                Set fso = Nothing

                '----------------------------------------
                ' Actualizar DB
                '----------------------------------------
                actualizarfoto a, nombreNuevo

                '----------------------------------------
                ' Redirección original
                '----------------------------------------
                If Request.Cookies("edit_origen") = 1 Then
                    Response.Redirect "lista_total.asp?tv=" & Request.Cookies("edit_verPantalla") & "&e=" & Request.Cookies("edit_estatusAnuncio") & "&op=" & Request.Cookies("edit_ordenadoPor")
                Else
                    Response.Redirect "lista.asp?tv=" & Request.Cookies("edit_verPantalla") & "&e=" & Request.Cookies("edit_estatusAnuncio") & "&op=" & Request.Cookies("edit_ordenadoPor")
                End If

            Else
        %>

            <!-- Form de prueba -->
            <form method="POST" enctype="multipart/form-data">
                <input type="file" name="file"><br><br>
                <input type="submit" value="Subir">
            </form>

        <%
            End If
        %>
    </body>
</html>