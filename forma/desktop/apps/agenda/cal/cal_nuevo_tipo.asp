<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function NuevoCodigo()
                dim cc, tt, sqlString

                sqlString = "SELECT MAX(CAST(Codigo AS Numeric(3, 0))) + 1 AS nCodigo " & _
                              "FROM cal_Calendarios " & _
                             "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                               "AND  (Isnumeric(Codigo) = 1);"
                
                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.execute(sqlString)

                    if (tt.bof or tt.eof) then
                        NuevoCodigo = "000"
                    else
                        NuevoCodigo = RIGHT("000" & tt("nCodigo"), 3)
                    end if

                tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function
        %>
    </head>

    <body>
        <%
            dim cc, sqlString
            dim codigo, Nombre, Usuario

            Usuario = Request.Cookies("Usuario")
            nombre = Request.Form("nuevoNombre")

            sqlString = "INSERT INTO cal_Calendarios(Usuario, Codigo, Nombre, PorDefecto, DeSistema, ColorFont, Seleccionado) " & _
                        "VALUES ('" & Usuario & "', '" & NuevoCodigo() & "', '" & nombre & "', 0, 0, 'rgb(0,0,0)', 1);"

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")
                cc.execute(sqlString)
            cc.close: set cc = nothing

            response.redirect "cal_tipos.asp"
        %>    
    </body>
</html>