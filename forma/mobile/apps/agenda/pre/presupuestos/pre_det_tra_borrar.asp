<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            sub Actualizar_Incrementos(Presupuesto)
                dim cc, sqlString

                sqlString = "exec dbo.pre_Presupuestos_Incrementos '" & Request.Cookies("Usuario") & "', '" & Presupuesto & "'"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                    cc.execute sqlString  
                cc.close: set cc = nothing            
            end sub               
        %>
    </head>

    <body>
        <%
            dim con, pre, editor, llave

            pre = request.QueryString("presupuesto")
            llave = request.QueryString("registro")
            
            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")
                con.Execute ("DELETE FROM pre_Presupuesto_Detalles WHERE Llave = " & llave & ";")
            con.close: set con = nothing

            '
            ' Actualizamos los Incrementos si es un Modelo...
            '
            Actualizar_Incrementos pre

            '
            ' Volvemos al Editor
            '
            response.redirect "pre_det_editar.asp"
        %>    
    </body>
</html>