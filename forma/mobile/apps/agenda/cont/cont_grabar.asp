<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function SecuenciaContacto()
                dim c, t, sqlString, usu, valor 
                usu = Request.Cookies("Usuario")

                sqlString = "select (MAX(CAST(Codigo as Numeric(12,0))) + 1) AS NuevoCodigo " & _
                            "from con_Contactos " & _
                            "where (Usuario = '" & usu & "') " & _
                            "and (Codigo <> '" & usu & "');"

                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                set t = c.execute(sqlString)

                if (t.bof or t.eof) then
                    valor = 1
                else
                    valor = t("NuevoCodigo")
                end if

                SecuenciaContacto = RIGHT("000000000000" & valor, 12)

                t.close: set t = nothing
                c.close: set c = nothing
            end function

            sub PrimerTelefono(telefono, contacto)
                dim c, t, sqlString, usu, valor 
                usu = Request.Cookies("Usuario")

                if telefono <> "" then
                    sqlString = "INSERT INTO con_Contactos_Telefonos(Usuario, Codigo, Telefono, Tipo) " & _
                                "VALUES('" & usu & "', '" & contacto & "', '" & telefono & "', 0);"

                    set c = Server.CreateObject("ADODB.Connection")
                    c.open Application("Conn")
                        c.execute(sqlString)
                    c.close: set c = nothing            
                end if
            end sub  

            sub ActualizarPrimerTelefono(telefono, contacto)
                dim c, t, sqlString, sqlBorrar, usu, valor 
                usu = Request.Cookies("Usuario")

                if telefono <> "" then
                    sqlBorrar = "DELETE FROM con_Contactos_Telefonos " & _
                                "WHERE (Usuario = '" & usu & "') " & _
                                "AND (Codigo = '" & contacto & "') " & _
                                "AND (Tipo = 0);"

                    sqlString = "INSERT INTO con_Contactos_Telefonos(Usuario, Codigo, Telefono, Tipo) " & _
                                "VALUES('" & usu & "', '" & contacto & "', '" & telefono & "', 0);"

                    set c = Server.CreateObject("ADODB.Connection")
                    c.open Application("Conn")
                        c.execute(sqlBorrar)
                        c.execute(sqlString)
                    c.close: set c = nothing            
                end if
            end sub              

            sub PrimerCategoria(tipo, contacto)
                dim c, t, sqlString, usu, valor, categ
                usu = Request.Cookies("Usuario")

                select case tipo
                    case "PE": categ = "principal"
                    case "ES": categ = "locales"
                    case "CU": categ = "cuenta"
                end select

                sqlString = "INSERT INTO con_Contactos_ConCategs(Usuario, Tipo, Categoria, Codigo)" & _
                            "VALUES('" & usu & "', '" & tipo & "', '" & categ & "', '" & contacto & "');"

                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    c.execute(sqlString)
                c.close: set c = nothing            
            end sub      

            Function Limpiar(Cadena)               
                Limpiar = Replace(Cadena,"'","´")
            end function
        %>
    </head>

    <body>
        <%
            '
            ' Grabar datos de un Contacto
            '
            dim cc, tt, sqlString, nuevo, ver, tipo, categ, orden1, orden2
            dim Usuario, Codigo, TipoContacto, PrimerNombre, SegundoNombre, PrimerApellido, SegundoApellido
            dim CorreoElectronico, FechaCumple, Empresa, TelefonoEmpresa, SitioWeb, Pais, Provincia, Ciudad, Direccion, Notas
            dim DeSistema, Visible, Arbol, Signo, telPrincipal

            Usuario = Request.Cookies("Usuario")
            Codigo = Request.Form("cod")
            TipoContacto = Request.Form("tipoContacto")
            PrimerNombre = limpiar(Request.Form("primerNombre"))
            SegundoNombre = limpiar(Request.Form("segundoNombre"))
            PrimerApellido = limpiar(Request.Form("primerApellido"))
            SegundoApellido = limpiar(Request.Form("segundoApellido"))
            CorreoElectronico = Request.Form("correoElectronico")
            FechaCumple = Request.Form("fechaCumple")

            Empresa = limpiar(Request.Form("empresa"))
            TelefonoEmpresa = Request.Form("telefonoEmpresa")
            SitioWeb = Request.Form("sitioWeb")
            Pais = limpiar(Request.Form("cboPais"))
            Provincia = limpiar(Request.Form("provincia"))
            Ciudad = limpiar(Request.Form("ciudad"))
            Direccion = limpiar(Request.Form("txtAreaDireccion"))

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")

                if Codigo = "nuevo" then
                    '
                    ' Si el Contacto es Nuevo, creamos un registro "en blanco" y lo preparamos
                    ' para actualizarlo con los datos iniciales. De otro modo, solo actualizamos
                    ' el registro existente
                    '
                    Codigo = SecuenciaContacto()
                    Notas = ""
                    telPrincipal = replace(Request.Form("tel"), "*", "+")

                    sqlString = "INSERT INTO con_Contactos(Usuario, Codigo, PrimerNombre) " & _
                                "VALUES('" & Usuario & "', '" & Codigo & "', '" & PrimerNombre & "');"
                    cc.execute(sqlString)

                    PrimerTelefono telPrincipal, Codigo
                    PrimerCategoria TipoContacto, Codigo
                end if

                sqlString = "UPDATE con_Contactos " & _
                               "SET PrimerNombre = '" & PrimerNombre & "'," & _
                                  " SegundoNombre = '" & SegundoNombre & "'," & _
                                  " PrimerApellido = '" & PrimerApellido & "'," & _
                                  " SegundoApellido = '" & SegundoApellido & "'," & _
                                  " CorreoElectronico = '" & CorreoElectronico & "'," & _
                                  " FechaCumple = '" & FechaCumple & "'," & _
                                  " Empresa = '" & Empresa & "'," & _
                                  " TelefonoEmpresa = '" & TelefonoEmpresa & "'," & _
                                  " SitioWeb = '" & SitioWeb & "'," & _
                                  " Pais = '" & Pais & "'," & _
                                  " Provincia = '" & Provincia & "'," & _
                                  " Ciudad = '" & Ciudad & "'," & _
                                  " Direccion = '" & Direccion & "'" & _
                             "WHERE (Usuario = '" & Usuario & "') " & _
                               "AND (Codigo = '" & Codigo & "');"

            cc.execute(sqlString)
            cc.close: set cc = nothing

            ActualizarPrimerTelefono Request.Form("tel"), Codigo

            response.redirect "lista.asp"
        %>    
    </body>
</html>