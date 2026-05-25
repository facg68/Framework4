<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Archivar Presupuesto" %>
        <title><%= PageTitle %></title>

		<%
            function EstatusPresupuesto(Presupuesto, Usuario)
                dim f, ft

                set f = Server.CreateObject("ADODB.Connection")
                f.open Application("Conn")

                set ft = f.execute("SELECT Estatus " & _
                                    "FROM pre_Presupuesto_Encabezado " & _
                                    "WHERE Usuario = '" & Usuario & "' " & _
                                    "AND Presupuesto = '" & Presupuesto & "';")

                EstatusPresupuesto = ft("Estatus")
                
                ft.close: set ft = nothing
                f.close: set f = nothing
            end function
		%>		        
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <%
            dim c, pre, usu, sqlString, EstatusActual

            pre = Request.QueryString("pre")
            usu = Request.Cookies("Usuario")

            EstatusActual = EstatusPresupuesto(pre, usu)

            If EstatusActual = 2 then 
                EstatusActual = 1
            else
                EstatusActual = 2
            end if

            sqlString = "UPDATE pre_Presupuesto_Encabezado " & _
                           "SET Estatus = " & EstatusActual & _
                        " WHERE Usuario = '" & usu & "' " & _
                           "AND Presupuesto = '" & pre & "';"

            set c = server.CreateObject("ADODB.Connection")
            c.open Application("Conn")
                c.execute (sqlString)
            c.close: set c = nothing

            response.redirect "../lista.asp"
        %>

        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>