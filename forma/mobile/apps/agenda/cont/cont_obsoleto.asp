<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function Estado(Contacto)
                dim cc, tt

                sqlString = "SELECT visible FROM con_Contactos " & _
                            "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                            "AND (Codigo = '" & Contacto &"');"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.execute(sqlString)

                if not (tt.bof or tt.eof) then
                    Estado = tt("Visible")
                end if

                tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function
        %>
    </head>

    <body>
        <%
            dim conn, sqlString, nEstatus
            dim t, c, v, o1, o2, con

            con = Request.QueryString("con")

            set conn = Server.CreateObject("ADODB.Connection")
            conn.open Application("Conn")

                select case Estado(con)
                    case 0 
                        nEstatus = 1
                    case 1
                        nEstatus = 0
                end select

                sqlString = "UPDATE con_Contactos " & _
                            "SET Visible = '" & nEstatus & "' " & _
                            "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                            "AND (Codigo = '" & con & "');"

                conn.execute(sqlString)

            conn.close: set conn = nothing
            response.redirect "lista.asp"
        %>
    </body>
</html>