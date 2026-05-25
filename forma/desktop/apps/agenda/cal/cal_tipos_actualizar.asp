<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            dim cc, tt, sqlString, comando
            dim valor, campo

            Usuario =  Request.Cookies("Usuario") 

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")
        %>
    </head>

    <body>
        <%
                sqlString = "SELECT Codigo FROM cal_Calendarios WHERE (Usuario = '" & Usuario & "');"

                set tt = cc.execute(sqlString)
                    if not (tt.bof or tt.eof) then
                        Do
                            campo = "ck_" & tt("Codigo")
                            valor = Request.Form(Campo)
                            comando = "UPDATE cal_Calendarios SET "
                                if valor = 1 then
                                    comando = comando & "Seleccionado = 1 "
                                else
                                    comando = comando & "Seleccionado = 0 "
                                end if 
                            comando = comando & "WHERE (Usuario = '" & Usuario & "') AND (Codigo = '" & tt("Codigo") & "');"

                            cc.execute(comando)
                            tt.MoveNext
                        Loop until tt.eof
                    end if
                tt.close: set tt = nothing
            cc.close: set cc = nothing

            response.redirect "cal_tipos.asp"
        %>
    </body>
</html>