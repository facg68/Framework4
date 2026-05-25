<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <!-- Upload libre -->
        <!-- #include virtual = "/core/includes/freeaspupload.asp" -->

        <%
            Response.Expires = -1
            Server.ScriptTimeout = 600
            Session.CodePage = 65001

            '----------------------------------------
            ' Resize con WIA
            '----------------------------------------
            sub ResizeGDI(originalPath, newPath, newWidth)
                On Error Resume Next

                dim img, width, height, newHeight
                dim ip, imgResized

                Set img = Server.CreateObject("WIA.ImageFile")
                img.LoadFile originalPath

                width = img.Width
                height = img.Height

                if width > 0 then
                    newHeight = (newWidth * height) / width
                else
                    newHeight = newWidth
                end if

                Set ip = Server.CreateObject("WIA.ImageProcess")
                ip.Filters.Add ip.FilterInfos("Scale").FilterID
                ip.Filters(1).Properties("MaximumWidth") = newWidth
                ip.Filters(1).Properties("MaximumHeight") = newHeight

                Set imgResized = ip.Apply(img)
                imgResized.SaveFile newPath

                set img = nothing
                set ip = nothing
                set imgResized = nothing

                If Err.Number <> 0 Then
                    Response.Write "Error en Resize: " & Err.Description & "<br>"
                    Err.Clear
                End If

                On Error Goto 0
            end sub
        %>
    </head>

    <body>
        <%
            Dim Upload, uploadsDir
            Dim fileKey, NombreOriginal
            Dim contacto, ver, tipo, categ, orden1, orden2
            Dim basePath, tempPath, finalPath, thumbPath
            Dim fso, vinculo

            '----------------------------------------
            ' Carpeta destino
            '----------------------------------------
            uploadsDir = Server.MapPath(lcase(Request.Cookies("usuPath")) & "/fotos")

            If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

                Set Upload = New FreeASPUpload
                Upload.Save(uploadsDir)

                '----------------------------------------
                ' Leer parámetros
                '----------------------------------------
                contacto = Upload.Form("contacto")
                ver      = Upload.Form("ver")
                tipo     = Upload.Form("tipo")
                categ    = Upload.Form("categ")
                orden1   = Upload.Form("orden1")
                orden2   = Upload.Form("orden2")

                If contacto = "" Then
                    Response.Write "Falta 'contacto'<br>"
                    Response.End
                End If

                '----------------------------------------
                ' Archivo subido
                '----------------------------------------
                If Upload.UploadedFiles.Count = 0 Then
                    Response.Write "No se subió archivo<br>"
                    Response.End
                End If

                For Each fileKey In Upload.UploadedFiles.Keys
                    NombreOriginal = Upload.UploadedFiles(fileKey).FileName
                Next

                '----------------------------------------
                ' Paths
                '----------------------------------------
                basePath   = uploadsDir
                tempPath   = basePath & "\" & NombreOriginal
                finalPath  = basePath & "\" & contacto & ".jpg"
                thumbPath  = basePath & "\" & contacto & "_s.jpg"

                '----------------------------------------
                ' Renombrar
                '----------------------------------------
                Set fso = Server.CreateObject("Scripting.FileSystemObject")

                If fso.FileExists(finalPath) Then fso.DeleteFile finalPath
                If fso.FileExists(thumbPath) Then fso.DeleteFile thumbPath

                fso.MoveFile tempPath, finalPath
                Set fso = Nothing

                '----------------------------------------
                ' Resize
                '----------------------------------------
                Call ResizeGDI(finalPath, thumbPath, 60)

                '----------------------------------------
                ' Redirección original
                '----------------------------------------
                vinculo = "cont_editar.asp?con=" & contacto & _
                        "&v=" & ver & _
                        "&t=" & tipo & _
                        "&c=" & categ & _
                        "&o1=" & orden1 & _
                        "&o2=" & orden2 & _
                        "&tt=1"

                Response.Redirect vinculo

            Else
        %>

            <!-- Form de prueba -->
            <form method="POST" enctype="multipart/form-data">
                <input type="text" name="contacto" placeholder="Contacto"><br><br>
                <input type="file" name="file"><br><br>
                <input type="submit" value="Subir">
            </form>

        <%
            End If
        %>
    </body>
</html>