<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 

<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <!-- Upload libre -->
        <!-- #include virtual = "/core/includes/freeaspupload.asp" -->

        <%
            Function Extension(NombreArchivo)
                Dim pos
                pos = InStrRev(NombreArchivo, ".")
                
                If pos > 0 Then
                    Extension = LCase(Mid(NombreArchivo, pos + 1))
                Else
                    Extension = ""
                End If
            End Function        
        %>
    </head>

    <body>    
        <%
            Dim p, NombreReal, Descripcion, Contacto, NuevoNombre
            Dim filesPath, filesExtension, fOriginal, fNuevo
            Dim Upload, fileKey

            '----------------------------------------
            ' Carpeta destino
            '----------------------------------------
            p = LCase(Request.Cookies("usuPath")) & "/adjuntos"
            filesPath = LCase(Server.MapPath(p))

            If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

                Set Upload = New FreeASPUpload
                Upload.Save(filesPath)

                '----------------------------------------
                ' Obtener archivo subido
                '----------------------------------------
                NombreReal = ""

                If Upload.UploadedFiles.Count > 0 Then
                    For Each fileKey In Upload.UploadedFiles.Keys
                        NombreReal = Upload.UploadedFiles(fileKey).FileName
                    Next
                End If

                '----------------------------------------
                ' Obtener datos del form
                '----------------------------------------
                Contacto     = Upload.Form("NuevoObjetoCodigoCont")
                Descripcion  = Upload.Form("NuevoObjeto")
                NuevoNombre  = Upload.Form("NuevoObjetoCSecuencia")

                '----------------------------------------
                ' Validaciones
                '----------------------------------------
                If Trim(NombreReal) = "" Then
                    Response.Write "ERROR: No se subió ningún archivo"
                    Response.End
                End If

                If Trim(NuevoNombre) = "" Then
                    Response.Write "ERROR: Nombre interno vacío"
                    Response.End
                End If

                '----------------------------------------
                ' Construcción de rutas
                '----------------------------------------
                filesExtension = Extension(NombreReal)

                fOriginal = filesPath & "\" & NombreReal
                fNuevo   = filesPath & "\" & NuevoNombre & "." & filesExtension

                '----------------------------------------
                ' Renombrar archivo
                '----------------------------------------
                Dim fso
                Set fso = CreateObject("Scripting.FileSystemObject")

                If fso.FileExists(fOriginal) Then

                    If fso.FileExists(fNuevo) Then
                        fso.DeleteFile fNuevo
                    End If

                    fso.MoveFile fOriginal, fNuevo

                Else
                    Response.Write "ERROR: Archivo original no encontrado<br>"
                    Response.Write fOriginal
                    Response.End
                End If

                Set fso = Nothing    

                '----------------------------------------
                ' Guardar en base de datos
                '----------------------------------------
                Dim sqlString, cc

                sqlString = "INSERT INTO con_Contactos_Adjuntos(Usuario, Codigo, Descripcion, Nombre, Extension) " & _
                            "VALUES('" & Request.Cookies("Usuario") & "', '" & Contacto & "', '" & Descripcion & "', '" & NuevoNombre & "', '" & filesExtension & "');"

                Set cc = Server.CreateObject("ADODB.Connection")
                cc.Open Application("Conn")
                    cc.Execute sqlString
                cc.Close
                Set cc = Nothing    

                '----------------------------------------
                ' Redirección
                '----------------------------------------
                Response.Redirect "cont_editar.asp?con=" & Contacto & "&tt=4"

            Else
        %>

            <!-- Form de prueba -->
            <form method="POST" enctype="multipart/form-data">
                <input type="text" name="NuevoObjetoCodigoCont" placeholder="Contacto"><br><br>
                <input type="text" name="NuevoObjeto" placeholder="Descripción"><br><br>
                <input type="text" name="NuevoObjetoCSecuencia" placeholder="Nombre interno"><br><br>
                <input type="file" name="file"><br><br>
                <input type="submit" value="Subir">
            </form>

        <%
            End If
        %>
    </body>
</html>