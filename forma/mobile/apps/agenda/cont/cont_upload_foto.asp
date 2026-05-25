<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <%
            sub Resize(contacto)
                imagePath =  lcase(Request.Cookies("usuPath")) & "/fotos"
                image_name = contacto & ".jpg" 
                new_image_filename = contacto & "_s.jpg" 

                Set Jpeg = Server.CreateObject("Persits.Jpeg") 

                Path = Server.MapPath(imagePath & "/" & image_name) 
                Jpeg.Open Path 

                Jpeg.Width = 60
                Jpeg.Height = cInt((60 * Jpeg.OriginalHeight) / Jpeg.OriginalWidth)

                Jpeg.Save Server.MapPath(imagePath & "/" & new_image_filename)         

                set Jpeg = nothing
            end sub
        %>           
    </head>

    <body>
        <% 
            dim p, separador, contacto, ver, tipo, categ 
            dim orden1, orden2, vinculo, fOriginal, fNuevo
            dim tam, donde, fso

            p = lcase(Request.Cookies("usuPath")) & "/fotos"

            '
            ' Subimos la foto a la carpeta...
            '

            Set Upload = Server.CreateObject("Persits.Upload") 
            count = upload.savevirtual(p)

            response.write "Files:<BR>"

            For Each File in Upload.Files
                '
                ' Buscamos el nombre del archivo
                ' Aunque esta dentro de un Loop, solo hay
                ' un solo item            
                '
                fOriginal = cStr(File.Path)
            Next

            '
            ' Separamos el Path del nombre del archivo original           
            '

            tam = len(fOriginal)
            donde = -1
            separador = 0            

            for k = tam to 1 step -1
                if mid(fOriginal, k, 1) = "\" then
                    if (donde = -1) then
                        donde = k
                    end if
                end if
            next

            if donde <> -1 then
                separador = donde
            end if

            '
            ' Cargamos los parámetros          
            '

            For Each Item in Upload.Form
                select case Item.Name
                    case "contacto": contacto = Item.Value
                end select
            Next

            '
            ' Copiamos el archivo original con el nombre
            ' que le corresponde y luego borramos el original
            '

            fNuevo =  left(fOriginal, separador) & contacto & ".jpg"

            Set fso = CreateObject("Scripting.FileSystemObject")
                fso.CopyFile fOriginal, fNuevo
                fso.DeleteFile fOriginal
            set fso = nothing    

            Resize contacto

            '
            ' Volvemos al formulario de edicion
            '

            vinculo = "cont_editar.asp?con=" & Contacto
            response.redirect vinculo
        %>  
    </body>
</html>