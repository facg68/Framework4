<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            Function NuevoCodigoTipo(Usuario)
                dim cc, tt, sqlString

                sqlString = "SELECT ISNULL(MAX(Codigo), 0) + 1 AS NuevaSec " & _
                              "FROM con_Contactos_Tipos " & _
                             "WHERE (Usuario = '" & Usuario & "') " & _
                               "AND (ISNUMERIC(Codigo) = 1);"
                
                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.execute(sqlString)
                    NuevoCodigoTipo = RIGHT("00" & tt("NuevaSec"), 2)
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

            sqlString = "INSERT INTO con_Contactos_Tipos(Usuario, Codigo, Nombre, Def, DeSistema) " & _
                        "VALUES ('" & usu & "', '" & NuevoCodigoTipo(usu) & "', '" & nombre & "', 0, 0);"

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")
                cc.execute(sqlString)
            cc.close: set cc = nothing

            response.redirect "cont_tipos.asp"
        %>    
    </body>
</html>