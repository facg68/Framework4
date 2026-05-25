<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            sub Actualizar_Incrementos(Presupuesto)
                dim cc, sqlString

                sqlString = "exec dbo.pre_Presupuestos_Incrementos '" & Request.Cookies("Usuario") & "', '" &  Presupuesto & "'"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                    cc.execute sqlString  
                cc.close: set cc = nothing            
            end sub        
        %>    
    </head>

    <body>
        <%
            Dim con, pre, sqlCommand
             
            pre = Request.QueryString("p")

            sqlCommand = "UPDATE pre_Presupuesto_Encabezado " & _
                            "SET Tipo = 'M', " & _
                               " Estatus = 1, " & _
                               " Cuantificable = 0 " & _
                         " WHERE Usuario = '" & request.Cookies("Usuario") & "' " & _
                           " AND Presupuesto = '" & pre & "';"

            set con = Server.CreateObject("ADODB.Connection")
            con.Open Application("Conn")
                con.Execute(sqlCommand)
            con.close: set con = nothing

            '
            ' Actualizamos los Incrementos en el Modelo...
            '
            Actualizar_Incrementos pre
           
            response.redirect "../lista.asp"
        %>    
    </body>
</html>