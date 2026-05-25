<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function QueVisibilidad(Usuario, Paquete, Objeto)
                dim con, tt, sqlString, visibilidad
                
                sqlString = "SELECT Visible " & _
                            "FROM discos_Objetos " & _
                        " WHERE (Usuario  = '" & usuario & "') " & _
                            "AND (Paquete = '" & Paquete & "') " & _
                            "AND (Objeto  = '" & objeto & "');"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.execute(sqlString)
                    QueVisibilidad = tt("Visible")
                tt.close: set tt = nothing
                cc.close: set cc = nothing            
            end function
        %>
    </head>

    <body>
        <%
            dim cc, tt, sqlString, visibilidad
            dim usuario, paquete, objetos, editor

            usuario = Request.Cookies("Usuario")
            Paquete = Request.QueryString("p")
            Objeto = Request.QueryString("o")

            visibilidad = QueVisibilidad(Usuario, Paquete, Objeto)
            if visibilidad = 1 then 
                visibilidad = 0
            else
                visibilidad = 1
            end if

            sqlString = "UPDATE discos_Objetos " & _
                        "SET Visible = " & visibilidad & _
                       " WHERE (Usuario  = '" & usuario & "') " & _
                        "AND (Paquete = '" & Paquete & "') " & _
                        "AND (Objeto  = '" & objeto & "');"

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")
                cc.execute(sqlString)
            cc.close: set cc = nothing

            response.redirect "editar.asp?m=" & paquete
        %>
    </body>
</html>