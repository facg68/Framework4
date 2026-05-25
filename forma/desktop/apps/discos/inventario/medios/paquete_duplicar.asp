<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <%
        dim cc, tt, sqlString, secuencia

        set cc = Server.CreateObject("ADODB.ConnectioN")
        cc.open Application("Conn")
    %>

    <head>
        <%
            Function SecuenciaPaquete()
                sqlString = "select ISNULL(MAX(CAST(RIGHT(paquete, 6) AS numeric(6,0))), 0) AS Valor " & _
                            "from discos_Paquetes " & _
                            "where Usuario ='" & Request.Cookies("Usuario") & "';"

                set tt = cc.execute(sqlString)
                    if not (tt.bof or tt.eof) then
                        secuencia = cDbl(tt("Valor")) + 1
                        SecuenciaPaquete = "PK" & RIGHT("000000" & secuencia, 6)
                    end if
                tt.close: set tt = nothing
            end function

            Function SecuenciaObjeto(Editor)
                sqlString = "select ISNULL(MAX(CAST(right(Objeto, 5) AS NUMERIC(5,0))),0) AS Valor " & _
                            "from discos_Objetos " & _
                            "where (usuario ='" & Request.Cookies("Usuario") & "') " & _
                            "and (Objeto Like '" & Editor & "%');"

                set tt = cc.execute(sqlString)
                    if not (tt.bof or tt.eof) then
                        secuencia = cDbl(tt("Valor")) + 1
                        SecuenciaObjeto = Editor & RIGHT("00000" & secuencia, 5)
                    end if
                tt.close: set tt = nothing
            end function

            sub CopiarPaquete(Usuario, Paquete, nPaquete)
                sqlString = "INSERT INTO discos_Paquetes(Usuario, Paquete, ACompra, AEdicion, Titulo, Precio, Tienda, Casa, Descripcion, VerComo, Carpeta) " & _
                            "SELECT Usuario, '" & nPaquete & "', ACompra, AEdicion, Titulo, Precio, Tienda, Casa, Descripcion, VerComo, Carpeta " & _
                            "FROM discos_Paquetes " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Paquete = '" & Paquete & "');"
                
                cc.execute(sqlString)            
            end sub

            sub CopiarObjeto(Usuario, Paquete, Objeto, nPaquete, nObjeto)
                sqlString = "INSERT INTO discos_Objetos(Usuario, Paquete, Objeto, AEdicion, Titulo, TituloOriginal, IdiomaMusica, InDirAu, DuraPag, Forma, Clasificacion, Tipo, FormatoPantalla, Recuento, Es3D, PlatOs, CopiaDigital, Descripcion, Editor, Visible) " & _
                            "SELECT Usuario, '" & nPaquete & "', '" & nObjeto & "', AEdicion, Titulo, TituloOriginal, IdiomaMusica, InDirAu, DuraPag, Forma, Clasificacion, Tipo, FormatoPantalla, Recuento, Es3D, PlatOs, CopiaDigital, Descripcion, Editor, Visible " & _
                            "FROM discos_Objetos " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Paquete = '" & Paquete & "') " & _
                            "AND (Objeto = '" & Objeto & "');"
                
                cc.execute(sqlString)
            end sub

            sub CopiarDetalles(Usuario, Paquete, Objeto, nPaquete, nObjeto)
                sqlString = "INSERT INTO discos_Objetos_Detalle(Usuario, Paquete, Objeto, Titulo, NumSerieLlave, Exito, Lado) " & _
                            "SELECT Usuario, '" & nPaquete & "', '" & nObjeto & "', Titulo, NumSerieLlave, Exito, Lado " & _
                            "FROM discos_Objetos_Detalle " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Paquete = '" & Paquete & "') " & _
                            "AND (Objeto = '" & Objeto & "');"
                
                cc.execute(sqlString)            
            end sub

            sub CopiarProtagonistas(Usuario, Paquete, Objeto, nPaquete, nObjeto)
                sqlString = "INSERT INTO discos_Objetos_Protagonistas(Usuario, Paquete, Objeto, Protagonista) " & _
                            "SELECT Usuario, '" & nPaquete &"', '" & nObjeto & "', Protagonista " & _
                            "FROM discos_Objetos_Protagonistas " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Paquete = '" & Paquete & "') " & _
                            "AND (Objeto = '" & Objeto & "');"
                
                cc.execute(sqlString)              
            end sub

            sub CopiarIdiomas(Usuario, Paquete, Objeto, nPaquete, nObjeto)
                sqlString = "INSERT INTO discos_Objetos_Idiomas(Usuario, Paquete, Objeto, Idioma, Audio, SubTitulos) " & _
                            "SELECT Usuario, '" & nPaquete & "', '" & nObjeto & "', Idioma, Audio, SubTitulos " & _
                            "FROM discos_Objetos_Idiomas " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Paquete = '" & Paquete & "') " & _
                            "AND (Objeto = '" & Objeto & "');"
                
                cc.execute(sqlString)              
            end sub

            sub CopiarImagenObjeto(Usuario, Objeto, nObjeto)
                On Error resume next

                dim iOriginal, iOriginal_s, iNueva, iNueva_s
                dim path

                path = Server.MapPath("\") & "\perfiles\" & lcase(Usuario) & "\medios\"

                iOriginal = path & Objeto & ".jpg"
                iOriginal_s = path & Objeto & "_s.jpg"
                iNueva = path & nObjeto & ".jpg"
                iNueva_s = path & nObjeto & "_s.jpg"

                Set fso = CreateObject("Scripting.FileSystemObject")
                    fso.CopyFile iOriginal, iNueva
                    fso.CopyFile iOriginal_s, iNueva_s
                set fso = nothing               
            end sub   

            sub CopiarImagenPaquete(Usuario, Paquete, nPaquete)
                On Error resume next

                dim iOriginal, iOriginal_s, iNueva, iNueva_s
                dim path

                path = Server.MapPath("\") & "\perfiles\" & lcase(Usuario) & "\medios\"

                iOriginal = path & Paquete & ".jpg"
                iOriginal_s = path & Paquete & "_s.jpg"
                iNueva = path & nPaquete & ".jpg"
                iNueva_s = path & nPaquete & "_s.jpg"

                Set fso = CreateObject("Scripting.FileSystemObject")
                    fso.CopyFile iOriginal, iNueva
                    fso.CopyFile iOriginal_s, iNueva_s
                set fso = nothing               
            end sub                       
        %>
    </head>

    <body>
        <%
            dim Usuario, Paquete, nPaquete, nObjeto, lista

            Usuario = Request.Cookies("Usuario")
            Paquete = Request.QueryString("p")

            nPaquete = SecuenciaPaquete()
            CopiarPaquete Usuario, Paquete, nPaquete

            set lista = cc.execute("SELECT Objeto, Editor FROM discos_Objetos WHERE Usuario = '" & Usuario & "' AND Paquete = '" & Paquete & "';")       
                if not (lista.bof or lista.eof) then
                    Do
                        Objeto = lista("Objeto")
                        Editor = lista("Editor")

                        nObjeto = SecuenciaObjeto(Editor)

                        CopiarObjeto Usuario, Paquete, Objeto, nPaquete, nObjeto
                        CopiarDetalles Usuario, Paquete, Objeto, nPaquete, nObjeto
                        CopiarProtagonistas Usuario, Paquete, Objeto, nPaquete, nObjeto
                        CopiarIdiomas Usuario, Paquete, Objeto, nPaquete, nObjeto
                        CopiarImagenObjeto Usuario, Objeto, nObjeto

                        lista.MoveNext
                    Loop Until lista.eof
                end if
            lista.close: set lista = nothing

            CopiarImagenPaquete Usuario, Paquete, nPaquete

            response.redirect "lista.asp"
        %>
    </body>

    <%
        cc.close: set cc = nothing
    %>
</html>