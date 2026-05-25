<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            '
            ' Variables Global
            '

            dim cc, tt, t, sqlString, sqlRegistro, Usuario
            dim paquete, objeto, editor, sec, tDet
            dim puntero, exito, numSerie
            dim protagonista, Lado, Audio, SubTitulos
            dim AEdicion, Titulo, TituloOriginal, IdiomaMusica, InDirAu
            dim DuraPag, Forma, Clasificacion, Tipo, FormatoPantalla, Recuento, Es3D, PlatOS
            dim CopiaDigital, Descripcion, Visible    

            Usuario = Request.Cookies("Usuario")     
            Paquete = Request.Form("Paquete")
            Objeto = Request.Form("Objeto")     
            AEdicion = Request.Form("AEdicion")
            Titulo = Limpiar(Request.Form("Titulo"))
            InDirAu = Limpiar(Request.Form("InDirAu"))
            Tipo = Request.Form("Tipo")
            Forma = Request.Form("Forma")
            Descripcion = Limpiar(Request.Form("Descripcion"))
            CopiaDigital = Request.Form("CopiaDigital")
            Editor = Request.Form("Editor")

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")

            '
            ' Funciones y Procedimientos
            '

            function EsMultilado(Usuario, Paquete, Objeto)
                sqlString = "SELECT f.Multilados " & _
                              "FROM discos_Objetos AS o " & _
                        "INNER JOIN discos_Formas AS f " & _
                                "ON o.Usuario = f.Usuario " & _
                               "AND o.Forma = f.Forma " & _
                            "WHERE (o.Usuario = '" & Usuario & "') " & _
                            "AND (o.Paquete = '" & Paquete & "') " & _
                            "AND (o.Objeto = '" & Objeto & "') " & _
                        "ORDER BY Titulo;"
                
                set t = cc.execute(sqlString) 
                
                    EsMultilado = t("Multilados")

                t.close: set t = nothing
            end function

            function Limpiar(valor)
                Limpiar = Replace(valor, "'", "´")
            end function    

            sub objetos_actualizar_detalles(Usuario, paquete, objeto, editor)
                sqlString = "SELECT Secuencia " & _
                            "FROM discos_Objetos_Detalle " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Paquete = '" & paquete & "') " & _
                            "AND (Objeto = '" & objeto & "') " & _
                        "ORDER BY Titulo;"
               
                set tDet = cc.execute(sqlString)

                if not (tDet.bof or tDet.eof) then
                    puntero = 0

                    Do
                        sec = tDet("Secuencia")
                        puntero = puntero + 1

                        select case editor
                            case "DM"
                                if EsMultilado(Usuario, paquete, objeto) = 1 then
                                    lado = Request.Form("DM01_FORM_Lado_" & puntero)
                                    exito = Request.Form("DM01_FORM_Exito_" & puntero)
                                    titulo = Limpiar(Request.Form("DM01_FORM_Titulo_" & puntero))

                                    sqlString = "UPDATE discos_Objetos_Detalle " & _
                                                "SET Lado = '" & lado & "', " & _
                                                    " Exito = " & exito & ", " & _
                                                    " Titulo = '" & titulo & "' " & _
                                            " WHERE Secuencia = " & sec & ";"

                                else
                                    exito = Request.Form("DM01_FORM_Exito_" & puntero)
                                    titulo = Limpiar(Request.Form("DM01_FORM_Titulo_" & puntero))

                                    sqlString = "UPDATE discos_Objetos_Detalle " & _
                                                "SET Exito = " & exito & ", " & _
                                                    " Titulo = '" & titulo & "' " & _
                                            " WHERE Secuencia = " & sec & ";"  
                                end if

                            case "VM"
                                exito = Request.Form("DM01_FORM_Exito_" & puntero)
                                titulo = Limpiar(Request.Form("DM01_FORM_Titulo_" & puntero))

                                sqlString = "UPDATE discos_Objetos_Detalle " & _
                                            "SET Exito = " & exito & ", " & _
                                                " Titulo = '" & titulo & "' " & _
                                        " WHERE Secuencia = " & sec & ";"

                            case "LI", "PE"
                                titulo = Limpiar(Request.Form("DM01_FORM_Titulo_" & puntero))

                                sqlString = "UPDATE discos_Objetos_Detalle " & _
                                            "SET Titulo = '" & titulo & "' " & _
                                        " WHERE Secuencia = " & sec & ";"       
                            
                            case "JU", "SO"
                                titulo = Limpiar(Request.Form("DM01_FORM_Titulo_" & puntero))
                                numSerie = Limpiar(Request.Form("DM01_FORM_NumSerie_" & puntero))                   

                                sqlString = "UPDATE discos_Objetos_Detalle " & _
                                            "SET Titulo = '" & titulo & "', " & _
                                            " NumSerieLlave = '" & numSerie & "' " & _
                                        " WHERE Secuencia = " & sec & ";"                  

                        end select

                        cc.execute(sqlString)
                        tDet.MoveNext
                    Loop Until tDet.eof
                end if

                tDet.close: set tDet = nothing         
            end sub   

            sub objetos_actualizar_protagonistas(Usuario, paquete, objeto, editor)
                sqlString = "SELECT Secuencia " & _
                            "FROM discos_Objetos_Protagonistas " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Paquete = '" & Paquete & "') " & _
                            "AND (Objeto = '" & Objeto & "') " & _
                        "ORDER BY Protagonista;"
                
                set t = cc.execute(sqlString)

                if not (t.bof or t.eof) then
                    puntero = 0

                    Do
                        puntero = puntero + 1
                        protagonista = Limpiar(Request.Form("DM01_FORM_Protagonista_" & puntero))

                        sqlString = "UPDATE discos_Objetos_Protagonistas " & _
                                    "SET Protagonista = '" & protagonista & "' " & _
                                " WHERE Secuencia = " & t("Secuencia") & ";"  

                        cc.execute(sqlString)
                        t.MoveNext
                    Loop Until t.eof
                end if

                t.close: set t = nothing       
            end sub      

            sub objetos_actualizar_idiomas(Usuario, paquete, objeto, editor)   
                sqlString = "SELECT oi.Idioma AS CodigoIdioma, oi.Audio, oi.Subtitulos, i.Nombre AS Idioma, oi.Secuencia " & _
                            "FROM discos_Objetos_Idiomas AS oi " & _
                        "INNER JOIN discos_Idiomas AS i " & _
                                "ON oi.Usuario = i.Usuario " & _
                            "AND oi.Idioma = i.Codigo " & _
                            "WHERE (oi.Usuario = '" & Usuario & "') " & _
                            "AND (oi.Paquete = '" & Paquete & "') " & _
                            "AND (oi.Objeto = '" & Objeto & "') " & _
                        "ORDER BY Idioma;"
                
                set t = cc.execute(sqlString)

                if not (t.bof or t.eof) then
                    puntero = 0

                    Do
                        puntero = puntero + 1
                        Audio = Request.Form("DM07_Audio_" & puntero)
                        SubTitulos = Request.Form("DM07_SubTitulos_" & puntero)

                        sqlString = "UPDATE discos_Objetos_Idiomas " & _
                                    "SET Audio = " & Audio & ", " & _
                                        " SubTitulos = " & SubTitulos & " " & _
                                " WHERE Secuencia = " & t("Secuencia") & ";"  

                        cc.execute(sqlString)
                        t.MoveNext
                    Loop Until t.eof
                end if

                t.close: set t = nothing        
            end sub             
        %>
    </head>

    <body>
        <%
            '
            ' Main()
            '

            Select Case Editor
                Case "DM"
                    Recuento = Request.Form("Recuento")
                    IdiomaMusica = Request.Form("IdiomaMusica")

                Case "VM"
                    Recuento = Request.Form("Recuento")
                    DuraPag = Request.Form("DuraPag")
                    Es3D = Request.Form("Es3D")
                    FormatoPantalla = Request.Form("FormatoPantalla")
                    Clasificacion = Request.Form("Clasificacion")
                    IdiomaMusica = Request.Form("IdiomaMusica")        

                Case "PE"
                    TituloOriginal = Limpiar(Request.Form("TituloOriginal"))
                    DuraPag = Request.Form("DuraPag")    
                    Es3D = Request.Form("Es3D")    
                    FormatoPantalla = Request.Form("FormatoPantalla")
                    Clasificacion = Request.Form("Clasificacion")        

                Case "JU", "SO"
                    Clasificacion = Request.Form("Clasificacion")        
                    PlatOS = Request.Form("PlatOS")   

                Case "LI"
                    DuraPag = Request.Form("DuraPag")        
                    
            End Select

            if Editor <> "HW" then
                sqlRegistro = "UPDATE discos_Objetos " & _
                                "SET AEdicion = " & AEdicion & ", " & _
                                    " Titulo = '" & Titulo & "', " & _
                                    " InDirAu = '" & InDirAu  & "', " & _
                                    " Tipo = '" & Tipo & "', " & _
                                    " Forma = '" & Forma & "', " & _
                                    " Descripcion = '" & Descripcion & "', " 

                if CopiaDigital = "" then 
                    sqlRegistro = sqlRegistro & " CopiaDigital = NULL, "
                else
                    sqlRegistro = sqlRegistro & " CopiaDigital = '" & CopiaDigital & "', "
                end if

                Select Case Editor
                    Case "DM"
                        sqlRegistro = sqlRegistro & " Recuento = " & Recuento & ", " & _
                                                    " IdiomaMusica = '" & IdiomaMusica & "' " 
                    Case "VM"
                        sqlRegistro = sqlRegistro & " Recuento = " & Recuento & ", " & _
                                                    " IdiomaMusica = '" & IdiomaMusica & "', " & _
                                                    " DuraPag = " & DuraPag & ", " & _
                                                    " Es3D = " & Es3D & ", " & _
                                                    " FormatoPantalla = '" & FormatoPantalla & "', " & _
                                                    " Clasificacion = '" & Clasificacion & "' " 

                    Case "PE"
                        sqlRegistro = sqlRegistro & " DuraPag = " & DuraPag & ", " & _
                                                    " Es3D = " & Es3D & ", " & _
                                                    " FormatoPantalla = '" & FormatoPantalla & "', " & _
                                                    " Clasificacion = '" & Clasificacion & "', " & _
                                                    " TituloOriginal = '" & TituloOriginal & "' "

                    Case "JU"
                        sqlRegistro = sqlRegistro & " Clasificacion = '" & Clasificacion & "', " & _
                                                    " PlatOS = '" & PlatOS & "' "
                    Case "SO"
                        sqlRegistro = sqlRegistro & " Clasificacion = '-', " & _
                                                    " PlatOS = '" & PlatOS & "' "                                                

                    Case "LI"
                        sqlRegistro = sqlRegistro & " DuraPag = " & DuraPag & " " 
                End Select                                      
            else
                sqlRegistro = "UPDATE discos_Objetos " & _
                                "SET AEdicion = " & AEdicion & ", " & _
                                    " Titulo = '" & Titulo & "', " & _
                                    " InDirAu = NULL, " & _
                                    " Tipo = '" & Tipo & "', " & _
                                    " Forma = '" & Forma & "', " & _
                                    " Descripcion = '" & Descripcion & "' "             
            end if

            sqlRegistro = sqlRegistro & "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                        "AND (Paquete = '" & Paquete & "') " & _
                                        "AND (Objeto = '" & Objeto & "');"
            
            cc.execute(sqlRegistro)

            '
            ' Actualizamos los detalles, idiomas y protagonistas...
            ' 
            select case Editor
                Case "DM", "VM", "LI", "SO", "JU"
                    objetos_actualizar_detalles Usuario, Paquete, Objeto, Editor
                Case "PE"
                    objetos_actualizar_detalles Usuario, Paquete, Objeto, Editor
                    objetos_actualizar_protagonistas Usuario, Paquete, Objeto, Editor
                    objetos_actualizar_idiomas Usuario, Paquete, Objeto, Editor
            end select

            cc.close: set cc = nothing
            response.redirect "editar.asp?m=" & Paquete
        %>