<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <%
        dim cc, tt, sqlString

        set cc = Server.CreateObject("ADODB.ConnectioN")
        cc.open Application("Conn")
    %>

    <head>
        <%
            Function SecuenciaObjeto(Editor)
                dim secuencia

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

            sub CopiarObjeto(Usuario, Paquete, Objeto, nObjeto)
                sqlString = "INSERT INTO discos_Objetos(Usuario, Paquete, Objeto, AEdicion, Titulo, TituloOriginal, IdiomaMusica, InDirAu, DuraPag, Forma, Clasificacion, Tipo, FormatoPantalla, Recuento, Es3D, PlatOs, CopiaDigital, Descripcion, Editor, Visible) " & _
                            "SELECT Usuario, Paquete, '" & nObjeto & "', AEdicion, Titulo, TituloOriginal, IdiomaMusica, InDirAu, DuraPag, Forma, Clasificacion, Tipo, FormatoPantalla, Recuento, Es3D, PlatOs, CopiaDigital, Descripcion, Editor, Visible " & _
                            "FROM discos_Objetos " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Paquete = '" & Paquete & "') " & _
                            "AND (Objeto = '" & Objeto & "');"
                
                cc.execute(sqlString)
            end sub

            sub CopiarDetalles(Usuario, Paquete, Objeto, nObjeto)
                sqlString = "INSERT INTO discos_Objetos_Detalle(Usuario, Paquete, Objeto, Titulo, NumSerieLlave, Exito, Lado) " & _
                            "SELECT Usuario, Paquete, '" & nObjeto & "', Titulo, NumSerieLlave, Exito, Lado " & _
                            "FROM discos_Objetos_Detalle " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Paquete = '" & Paquete & "') " & _
                            "AND (Objeto = '" & Objeto & "');"
                
                cc.execute(sqlString)            
            end sub

            sub CopiarProtagonistas(Usuario, Paquete, Objeto, nObjeto)
                sqlString = "INSERT INTO discos_Objetos_Protagonistas(Usuario, Paquete, Objeto, Protagonista) " & _
                            "SELECT Usuario, Paquete, '" & nObjeto & "', Protagonista " & _
                            "FROM discos_Objetos_Protagonistas " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Paquete = '" & Paquete & "') " & _
                            "AND (Objeto = '" & Objeto & "');"
                
                cc.execute(sqlString)              
            end sub

            sub CopiarIdiomas(Usuario, Paquete, Objeto, nObjeto)
                sqlString = "INSERT INTO discos_Objetos_Idiomas(Usuario, Paquete, Objeto, Idioma, Audio, SubTitulos) " & _
                            "SELECT Usuario, Paquete, '" & nObjeto & "', Idioma, Audio, SubTitulos " & _
                            "FROM discos_Objetos_Idiomas " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Paquete = '" & Paquete & "') " & _
                            "AND (Objeto = '" & Objeto & "');"
                
                cc.execute(sqlString)              
            end sub

            sub CopiarImagenes(Usuario, Objeto, nObjeto)
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
        %>
    </head>

    <body>
        <%
            dim Usuario, Paquete, Objeto, Editor, nObjeto

            Usuario = Request.Cookies("Usuario")

            Paquete = Request.QueryString("p")
            Objeto = Request.QueryString("o")
            Editor = Request.QueryString("e")

            nObjeto = SecuenciaObjeto(Editor)

            CopiarObjeto Usuario, Paquete, Objeto, nObjeto
            CopiarDetalles Usuario, Paquete, Objeto, nObjeto
            CopiarProtagonistas Usuario, Paquete, Objeto, nObjeto
            CopiarIdiomas Usuario, Paquete, Objeto, nObjeto
            CopiarImagenes Usuario, Objeto, nObjeto

            response.redirect "editar.asp?m=" & Paquete
        %>
    </body>

    <%
        cc.close: set cc = nothing
    %>
</html>