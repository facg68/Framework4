<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function LimpiarApostrofes(valor)
                LimpiarApostrofes = Replace(valor,"'","´")
            end function
        %>
    </head>

    <body>
        <%
            dim cc, tt, sqlString, num, nTitulo
            dim usuario, paquete, objetos, editor

            usuario = Request.Cookies("Usuario")
            Paquete = Request.QueryString("p")
            Objeto = Request.QueryString("o")
            Editor = Request.QueryString("e")

            sqlString = "select Secuencia, Titulo " & _
                        "from discos_Objetos_Detalle " & _
                        "where Usuario  = '" & usuario & "' " & _
                        "and Paquete = '" & Paquete & "' " & _
                        "and Objeto  = '" & objeto & "' " & _
                        "order by Secuencia;"

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")
            set tt = cc.execute(sqlString)

            if not (tt.bof or tt.eof) then
                num = 1
                Do
                    nTitulo = right("00" & num, 2) & " " & LimpiarApostrofes(tt("Titulo"))

                    sqlString = "UPDATE discos_Objetos_Detalle " & _
                                "SET Titulo = '" & nTitulo & "' " & _
                                "WHERE Secuencia = " & tt("Secuencia") & ";"

                    cc.execute(sqlString)
                    num = num + 1

                    tt.MoveNext
                loop until (tt.eof)
            end if

            tt.close: set tt = nothing
            cc.close: set cc = nothing

            response.redirect "editar_objeto.asp?p=" & paquete & "&o=" & objeto & "&e=" & Editor
        %>
    </body>
</html>