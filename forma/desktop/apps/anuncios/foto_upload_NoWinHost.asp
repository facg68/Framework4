<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/clsUpload/clsUpload.asp" -->    

        <%
            sub actualizarfoto(Anuncio, Nombre)
                dim con, sqlstring

                sqlstring = "UPDATE seg_anuncios SET NombreImagen = '" & Nombre & "' WHERE Secuencia = " & Anuncio & ";"

                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")

                    con.execute (sqlString)
                con.close: set con = nothing
            end sub

            function nom(original)        
                dim tam, k, donde

                tam = len(trim(original))

                if tam > 0 then
                    donde = -1

                    for k = tam to 1 step -1
                        if mid(original, k, 1) = "\" then
                            if donde = -1 then
                                donde = k
                            end if
                        end if
                    next

                    if donde = -1 then
                        nom = ""
                    else
                        nom = mid(original, (donde + 1), (tam - donde))
                    end if
                else
                    nom = ""
                end if
            end function

            function NombreImagen(secuencia)
                dim c, t, sqlStr

                sqlStr = "SELECT NombreImagen FROM seg_Anuncios " & _
                          "WHERE (Secuencia = " & Secuencia & ");"

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
                        path = Server.MapPath("\") & "\imagenes\anuncios\" 
                        imagen = path & Objeto

                        Set fso = CreateObject("Scripting.FileSystemObject")
                            fso.DeleteFile imagen
                        set fso = nothing      
                    end if   

                On Error Goto 0      
            end sub       

            Function Codigo(Anuncio)
                dim cc, tt, sqlstring

                sqlstring = "SELECT Codigo FROM seg_anuncios WHERE Secuencia = " & Anuncio & ";"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.execute(sqlstring)
                        Codigo = tt("Codigo")
                tt.close: set tt = nothing
                cc.close: set cc = nothing                
            end function  

            Function Extension(fOriginal)
                dim tam, donde, separador

                tam = len(fOriginal)
                donde = -1       

                for k = tam to 1 step -1
                    if mid(fOriginal, k, 1) = "." then
                        if (donde = -1) then
                            donde = k
                        end if
                    end if
                next

                if donde <> -1 then
                    Extension = right(fOriginal, (tam - donde))
                else
                    Extension = ""
                end if   
            End Function                    
        %>
    </head>

    <body>
        <% 
            dim p, separador, fOriginal, fNuevo
            dim tam, donde, fso, fotoAnuncio, a
            Dim objUpload, strFile, strPath

            Set objUpload = New clsUpload

            strFile = objUpload.Fields("file").FileName
            strPath = server.mappath("/forma/desktop/apps/anuncios/publicaciones") & "/" & strFile

            a = request.cookies("edit_anumcio")
            BorrarImagen a

            '
            ' Aqui se genera un nuevo nombre de archivo para que coincida con el
            ' filesystem. De esa forma no tendré problemas con caracteres especiales
            ' o nombres "tipo Windows" qua causarian conflictos con servidores de
            ' otro tipo
            '
            fNuevo = Codigo(a) & "." & Extension(strFile)       
            strPath2 = server.mappath("/forma/desktop/apps/anuncios/publicaciones") & "\" & fNuevo

            '
            ' Subimos la foto a la carpeta...
            '

            objUpload("file").SaveAs strPath2
            Set objUpload = Nothing
            
            actualizarfoto a, fNuevo

            '
            ' Volvemos al formulario de edicion
            '
            response.redirect "lista.asp"         
        %>
    </body>
</html>