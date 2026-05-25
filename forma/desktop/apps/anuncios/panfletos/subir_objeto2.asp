<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <!-- Upload libre -->
        <!-- #include virtual = "/core/includes/freeaspupload.asp" -->

        <%
            sub actualizarfoto(CU, Nombre)
                dim con, sqlstring

                sqlstring = "UPDATE seg_panfletos SET Objeto = '" & Nombre & "' WHERE CU = " & CU

                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")
                    con.execute sqlstring
                con.close: set con = nothing
            end sub

            function NombreImagen(CU)
                dim c, t, sqlStr

                sqlStr = "SELECT Objeto FROM seg_panfletos WHERE CU = " & CU

                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set t = c.execute(sqlStr)
                        NombreImagen = trim(t("Objeto"))
                    t.close: set t = nothing
                c.close: set c = nothing      
            end function

            sub BorrarImagen(CU)
                On Error resume next

                dim imagen, path, objeto

                Objeto = NombreImagen(CU)

                if len(trim(Objeto)) > 0 then
                    path = Server.MapPath("/forma/desktop/apps/anuncios/pdf") & "\" 
                    imagen = path & Objeto

                    Set fso = CreateObject("Scripting.FileSystemObject")
                        if fso.FileExists(imagen) then fso.DeleteFile imagen
                    set fso = nothing      
                end if   

                On Error Goto 0      
            end sub   
        %>
    </head>

    <body>
        <%
            Dim Upload, uploadsDir
            Dim fileKey, NombreOriginal
            Dim cu, basePath, tempPath, finalPath
            Dim fso

            cu = Request.Cookies("edit_Panfleto")

            '----------------------------------------
            ' Borrar archivo anterior
            '----------------------------------------
            BorrarImagen cu

            '----------------------------------------
            ' Carpeta destino
            '----------------------------------------
            uploadsDir = Server.MapPath("/forma/desktop/apps/anuncios/pdf")

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
                finalPath = basePath & "\" & cu & ".pdf"

                '----------------------------------------
                ' Renombrar a CU.pdf
                '----------------------------------------
                Set fso = CreateObject("Scripting.FileSystemObject")

                If fso.FileExists(finalPath) Then fso.DeleteFile finalPath
                fso.MoveFile tempPath, finalPath

                Set fso = Nothing

                '----------------------------------------
                ' Actualizar DB
                '----------------------------------------
                actualizarfoto cu, cu & ".pdf"

                '----------------------------------------
                ' Redirección original
                '----------------------------------------
                Response.Redirect "lista.asp?e=" & Request.Cookies("edit_estatusPanfleto") & "&op=" & Request.Cookies("edit_ordenadoPor")

            Else
        %>

            <!-- Form de prueba -->
            <form method="POST" enctype="multipart/form-data">
                <input type="file" name="file"><br><br>
                <input type="submit" value="Subir PDF">
            </form>

        <%
            End If
        %>
    </body>
</html>