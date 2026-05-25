<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>

        <title>Verificar Mi Ficha de Contacto</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            dim Con, t, sqlString, usuNombre, usuCorreo, usuFechaNacimiento, diaMes

            set Con = Server.CreateObject("ADODB.Connection")
            Con.open Application("Conn")

            thisSystem = "agenda"
            thisProcess = "agenda.0070"
            SysLockOut

            Function ExisteFicha()
                sqlString = "SELECT COUNT(*) AS Existe " & _
                              "FROM con_Contactos " & _
                             "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                               "AND (Codigo = '" & Request.Cookies("Usuario") & "');"

                set t = Con.execute(sqlString)
                    ExisteFicha = t("Existe")
                t.close: set t = nothing
            end function

            Sub CargarDatosUsuario(Usuario)
                dim dd, mm

                sqlString = "SELECT usuNombre, usuCorreo, usuFechaNacimiento " & _
                              "FROM dbo.seg_Usuarios " & _
                             "WHERE (usuCodigo = '" & Usuario & "');"

                set t = Con.execute(sqlString)
                    usuNombre = t("usuNombre")
                    usuCorreo = t("usuCorreo")
                    usuFechaNacimiento = t("usuFechaNacimiento")

                    if isnull(usuFechaNacimiento) = False then
                        dd = Day(t("usuFechaNacimiento"))
                        mm = Month(t("usuFechaNacimiento"))

                        diaMes = "'" & right("00" & dd, 2) & "/" & right("00" & mm, 2) & "'"
                    else
                        diaMes = "NULL"
                    end if
                t.close: set t = nothing            
            End Sub

            Sub CrearFicha(Usuario)
                CargarDatosUsuario Usuario

                sqlString = "INSERT INTO con_Contactos(Usuario, Codigo, PrimerNombre, CorreoElectronico, FechaCumple) " & _
                            "VALUES ('" & Usuario & "', '" & Usuario & "', '" & usuNombre & "', '" & usuCorreo & "', " & diaMes & ");"

                Con.execute(sqlString)                            
            End Sub
        %>            
    </head>

    <body>
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    

        <%
            if ExisteFicha() = 1 then
                response.redirect "cont_editar.asp?con=" & Request.Cookies("Usuario")
            else
                CrearFicha Request.Cookies("Usuario")
                response.redirect "cont_editar.asp?con=" & Request.Cookies("Usuario")
            end if
        %>

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->    
        <% Con.close: set Con = nothing %>
    </body>    
</html>