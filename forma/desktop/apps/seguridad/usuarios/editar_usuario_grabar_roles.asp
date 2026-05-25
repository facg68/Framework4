<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <%
        dim cc, tt, sqlString, usuario, sqlCommand

        set cc = Server.CreateObject("ADODB.Connection")
        cc.open Application("Conn")    
    %>

    <head>
        <%

            Sub ProcesarRoles(Usuario)
                dim roles, valor, comando, campo

                sqlCommand = "SELECT rolCodigo FROM seg_Roles;"

                set roles = cc.execute(sqlCommand)
                    if not (roles.bof or roles.eof) then
                        cc.execute("DELETE FROM seg_RolesUsuarios WHERE CodigoUsuario = '" & Usuario & "';")


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
            Codigo = Request.Form("codigo")   
            ProcesarRoles Codigo
        %>
    </body>

    <%
        cc.close: set cc = nothing
        Response.redirect "lista.asp?o=" & ordenadoPor    
    %>
</html>