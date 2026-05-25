<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
  <head>
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
      <%
        dim c, pre, usu, sqlString, EstatusActual

        pre = Request.QueryString("p")
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
  </body>
</html>

