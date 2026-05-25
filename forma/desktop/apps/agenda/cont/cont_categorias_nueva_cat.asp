<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            Function NuevoCodigoCateg(Usuario, Tipo)
                dim cc, tt, sqlString

                sqlString = "SELECT ISNULL(MAX(Codigo), 0) + 1 AS NuevaSec " & _
                               "FROM con_Contactos_Categorias " & _
                              "WHERE (Usuario = '" & Usuario & "') " & _
                                "AND (Tipo = '" & Tipo & "') " & _
                                "AND (ISNUMERIC(Codigo) = 1);"
                
                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                    set tt = cc.execute(sqlString)
                        NuevoCodigoCateg = RIGHT("00000000" & tt("NuevaSec"), 2)
                    tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function    
        %>
    </head>

    <body>
        <%
            dim cc, sqlString, Nombre, usu

            usu = Request.Cookies("Usuario")
            nombre = Request.Form("nuevoNombre")
            codTipo = Request.Form("codTipo")
            

            sqlString = "INSERT INTO con_Contactos_Categorias(Usuario, Tipo, Codigo, Nombre, PorDefecto, DeSistema) " & _
                        "VALUES ('" & usu & "', '" & codTipo & "', '" & NuevoCodigoCateg(usu, codTipo) & "', '" & nombre & "', 0, 0);"

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")
                cc.execute(sqlString)
            cc.close: set cc = nothing

            response.redirect "cont_categorias.asp?t=" & codTipo
        %>    
    </body>
</html>