<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <%
        dim cc, tt, sqlString, usuario, sqlCommand
        dim nombre, homepage, versaldos, IniciarSinEncabezado
        dim cargarSnippets, snippetsOpacidad, usuRandomWallpaper

        set cc = Server.CreateObject("ADODB.Connection")
        cc.open Application("Conn")    
    %>

    <head>
        <%
            function FechaServidor(FechaFormulario)
                dim d, m, a

                d = left(FechaFormulario, 2)
                m = mid(FechaFormulario, 4, 2)
                a = right(FechaFormulario, 4)

                FechaServidor = a & "-" & m & "-" & d
            end function

            function limpiar(cadena)
                limpiar = Replace(cadena,"'","´")    
            end function

            Sub CopiarFolder(Usuario)
                map = Server.MapPath("\perfiles\")
                fOriginal = map & "\defaults" 
                fNuevo = map & "\" & usuario

                Set fso = CreateObject("Scripting.FileSystemObject")
                    fso.CreateFolder(fnuevo)
                    fso.CopyFolder fOriginal, fNuevo, True
                set fso = nothing    
            end Sub

            Sub DiscosCarpetas(Usuario)
                sqlCommand = "INSERT INTO discos_Carpetas(Usuario, Codigo, Nombre, Descripcion, PorDefecto, DeSistema) " & _
                            "SELECT '" & Usuario & "' AS Usuario, Codigo, Nombre, Descripcion, PorDefecto, DeSistema " & _
                            "FROM dbo.discos_Carpetas " & _
                            "WHERE (Usuario = 'defaults');"
                cc.execute(sqlCommand)
            end Sub

            Sub DiscosTiendas(Usuario)
                sqlCommand = "INSERT INTO discos_Tiendas( Usuario, Codigo, Nombre, Contacto, SitioWeb, Correo, Tipo, Pais, Telefono1, Telefono2, Direccion, Notas, MediosDigitales, MediosFisicos, Musica, Video, Juegos, Software, Libros, Estatus) " & _
                            "SELECT '" & Usuario & "' AS Usuario, Codigo, Nombre, Contacto, SitioWeb, Correo, Tipo, Pais, Telefono1, Telefono2, Direccion, Notas, MediosDigitales, MediosFisicos, Musica, Video, Juegos, Software, Libros, Estatus " & _
                            "FROM dbo.discos_Tiendas " & _
                            "WHERE (Usuario = 'defaults');"
                cc.execute(sqlCommand)
            end Sub    

            Sub DiscosCasas(Usuario)
                sqlCommand = "INSERT INTO discos_Casas( Usuario, Codigo, Nombre, Musica, Video, Juegos, Software, Libros, Obsoleta) " & _
                            "SELECT '" & Usuario & "' AS Usuario, Codigo, Nombre, Musica, Video, Juegos, Software, Libros, Obsoleta " & _
                            "FROM dbo.discos_Casas " & _
                            "WHERE Usuario = 'defaults';"
                cc.execute(sqlCommand)
            end sub

            Sub DiscosFormatos(Usuario)
                sqlCommand = "INSERT INTO discos_formas( Usuario, Forma, Nombre, Multilados, Musica, Video, Juegos, Software, Libros, Icono_Forma, Estatus) " & _
                            "SELECT '" & Usuario & "' AS Usuario, Forma, Nombre, Multilados, Musica, Video, Juegos, Software, Libros, Icono_Forma, Estatus " & _
                            "FROM discos_Formas " & _
                            "WHERE Usuario = 'defaults';"
                cc.execute(sqlCommand)
            end sub    

            Sub DiscosGeneros(Usuario)
                sqlCommand = "INSERT INTO discos_Tipos (Usuario, Codigo, Nombre, Musica, Video, Juegos, Software, Libros) " & _
                            "SELECT '" & Usuario & "' AS Usuario, Codigo, Nombre, Musica, Video, Juegos, Software, Libros " & _
                            "FROM discos_Tipos " & _
                            "WHERE Usuario = 'defaults';"
                cc.execute(sqlCommand)
            end sub     

            Sub DiscosPlataformas(Usuario)
                sqlCommand = "INSERT INTO discos_Plataformas (Usuario, Codigo, Nombre, Software, Juegos, Obsoleta) " & _
                            "SELECT '" & Usuario & "' AS Usuario, Codigo, Nombre, Software, Juegos, Obsoleta " & _
                            "FROM discos_Plataformas " & _
                            "WHERE Usuario = 'defaults';"
                cc.execute(sqlCommand)
            end sub           

            Sub DiscosClasificaciones(Usuario)
                sqlCommand = "INSERT INTO discos_Clasificaciones(Usuario, Codigo, Nombre) " & _
                            "SELECT '" & Usuario & "' AS Usuario, Codigo, Nombre " & _
                            "FROM discos_Clasificaciones " & _
                            "WHERE Usuario = 'defaults';"
                cc.execute(sqlCommand)
            end sub    

            Sub DiscosFormatosPantalla(Usuario)
                sqlCommand = "INSERT INTO discos_FormatosPantalla (Usuario, Codigo, Nombre) " & _
                            "SELECT '" & Usuario & "' AS Usuario, Codigo, Nombre " & _
                            "FROM discos_FormatosPantalla " & _
                            "WHERE (Usuario = 'defaults');"
                cc.execute(sqlCommand)

                sqlCommand = "INSERT INTO discos_Graficas_Defaults (Usuario, Grafica, TipoAmo, SeleccionAmo, SeleccionMedio) " & _
                            "SELECT '" & Usuario & "' AS Usuario, Grafica, TipoAmo, SeleccionAmo, SeleccionMedio " & _
                            "FROM discos_Graficas_Defaults " & _
                            "WHERE (Usuario = 'defaults');"
                cc.execute(sqlCommand)

                sqlCommand = "INSERT INTO discos_Idiomas (Usuario, Codigo, Nombre) " & _
                            "SELECT '" & Usuario & "' AS Usuario, Codigo, Nombre " & _
                            "FROM discos_Idiomas " & _
                            "WHERE (Usuario = 'defaults')"
                cc.execute(sqlCommand)                
            end sub      

            Sub CalendarioTipos(Usuario)
                sqlCommand = "INSERT INTO cal_Calendarios (Usuario, Codigo, Nombre, PorDefecto, DeSistema, ColorFont, Seleccionado) " & _
                            "SELECT '" & Usuario & "' AS Usuario, Codigo, Nombre, PorDefecto, DeSistema, ColorFont, Seleccionado " & _
                            "FROM cal_Calendarios " & _
                            "WHERE (Usuario = 'defaults');"    

                cc.execute(sqlCommand)
            End Sub

            Sub ContactosTipos(Usuario)
                sqlCommand = "INSERT INTO con_Contactos_Tipos (Usuario, Codigo, Nombre, Def, DeSistema) " & _
                            "SELECT '" & Usuario & "' AS Usuario, Codigo, Nombre, Def, DeSistema " & _
                            "FROM con_Contactos_Tipos " & _
                            "WHERE (Usuario = 'defaults');"  

                cc.execute(sqlCommand)
            End Sub

            Sub ContactosCategorias(Usuario)
                sqlCommand = "INSERT INTO con_Contactos_Categorias (Usuario, Tipo, Codigo, Nombre, PorDefecto, DeSistema) " & _
                            "SELECT '" & Usuario & "' AS Usuario, Tipo, Codigo, Nombre, PorDefecto, DeSistema " & _
                            "FROM con_Contactos_Categorias " & _
                            "WHERE (Usuario = 'defaults');"

                cc.execute(sqlCommand)
            End Sub

            Sub Presupuestos(Usuario)
                sqlCommand = "INSERT INTO dbo.pre_Cuentas(Usuario, Codigo, Nombre, Categoria, Tipo, Anualidad, Monto, Contacto, LocalMonetario, MensajeDefault, TipoCuenta, Repetitiva, DeSistema, Grupo, Clase) " & _
                            "SELECT '" & Usuario & "' AS Usuario, Codigo, Nombre, Categoria, Tipo, Anualidad, Monto, Contacto, LocalMonetario, MensajeDefault, TipoCuenta, Repetitiva, DeSistema, Grupo, Clase " & _
                            "FROM dbo.pre_Cuentas " & _
                            "WHERE Usuario = 'defaults';"

                cc.execute(sqlCommand)
            End Sub        

            Sub ProcesarRoles(Usuario)
                dim roles, valor, comando, campo

                sqlCommand = "SELECT rolCodigo FROM seg_Roles;"

                set roles = cc.execute(sqlCommand)
                    if not (roles.bof or roles.eof) then
                        do
                            campo = roles("rolCodigo")
                            valor = Request.Form(campo)

                            if valor = 1 then
                                '-------------------------------------------------------------------'
                                ' El campo tiene "checked" por lo que asignamos este rol al usuario '
                                '-------------------------------------------------------------------'

                                comando = "INSERT INTO seg_RolesUsuarios(CodigoRol, CodigoUsuario, Activo) " & _
                                               "VALUES ('" & roles("rolCodigo") & "', '" & Usuario & "', 1);"
                                cc.execute(comando)
                            end if

                            roles.MoveNext
                        loop until(roles.eof)
                    end if
                roles.close: set roles = nothing
            end Sub
        %>
    </head>

    <body>
        <%
            Codigo = Request.Form("usucodigo")   
            nombre = limpiar(request.form("usuNombre"))

            sqlString = "INSERT INTO seg_Usuarios(usuCodigo, usuNombre) " & _
                                "VALUES ('" &  Codigo & "', '" &  nombre & "');"
            cc.execute(sqlString)    

            '----------------------------------------'
            ' Adiciones para la Extranet de Fabrizio '
            '----------------------------------------'

                CopiarFolder Codigo
                DiscosCarpetas Codigo
                DiscosTiendas Codigo
                DiscosCasas Codigo
                DiscosFormatos Codigo
                DiscosGeneros Codigo
                DiscosPlataformas Codigo
                DiscosClasificaciones Codigo
                DiscosFormatosPantalla Codigo
                CalendarioTipos Codigo
                ContactosTipos Codigo
                ContactosCategorias Codigo
                Presupuestos Codigo

                ProcesarRoles Codigo
        
            '------------------'
            ' Fin de Adiciones '
            '------------------'          
        %>
    </body>

    <%
        cc.close: set cc = nothing
        Response.redirect "lista.asp?o=" & ordenadoPor    
    %>
</html>