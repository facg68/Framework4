<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
  <head>
    <%
      Function Presupuesto_Secuencia(Usuario) 
        dim p, pt, NuevaSecuencia

        set p = Server.CreateObject("ADODB.Connection")
        p.open Application("Conn")

        set pt = p.execute("SELECT MAX(CAST(RIGHT(Presupuesto, 9) AS int)) AS Maximo " & _
                             "FROM dbo.pre_Presupuesto_Encabezado AS e " & _
                            "WHERE (LEFT(Presupuesto, 3) = 'PR-') " & _
                              "AND Usuario = '" & Usuario & "';")

          NuevaSecuencia = pt("Maximo") + 1                                     
          Presupuesto_Secuencia = "PR-" & RIGHT("000000000" & NuevaSecuencia, 9) 

        pt.close: set pt = nothing
        p.close: set p = nothing
      End Function

      Sub DuplicarDetalles(Usuario, Presupuesto, NuevoPresupuesto)
        dim f, sqlString, NuevoPre

        sqlString = "INSERT INTO pre_Presupuesto_Detalles(Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, Aplicado, Contacto, Nota, NotaPre, NotaDonde, Incremento, Archivado) " & _
                    "SELECT '" & NuevoPresupuesto & "', Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, Aplicado, Contacto, Nota, NotaPre, NotaDonde, Incremento, 0 " & _
                      "FROM pre_Presupuesto_Detalles " & _
                     "WHERE Usuario = '" & Usuario & "' " & _
                       "AND Presupuesto = '" & Presupuesto & "';"

        set f = Server.CreateObject("ADODB.Connection")
        f.open Application("Conn")
        f.execute(sqlString)
        f.close: set f = nothing
      end sub

      Sub DuplicarEncabezado(Usuario, Presupuesto, NuevoPresupuesto)
        dim f, sqlString

        sqlString = "INSERT INTO pre_Presupuesto_Encabezado(Presupuesto, Usuario, Tipo, Nombre, Desde, Hasta, SaldoFinal, MultiPrecio, MonedaOrigen, MonedaDestino, Estatus, Cuantificable) " & _
                    "SELECT '" & NuevoPresupuesto & "', Usuario, Tipo, Nombre, Desde, Hasta, SaldoFinal, MultiPrecio, MonedaOrigen, MonedaDestino, Estatus, Cuantificable " & _
                      "FROM pre_Presupuesto_Encabezado " & _
                     "WHERE Usuario = '" & Usuario & "' " & _
                       "AND Presupuesto = '" & Presupuesto & "';"

        set f = Server.CreateObject("ADODB.Connection")
        f.open Application("Conn")
        f.execute(sqlString)
        f.close: set f = nothing
      end sub
    %>
  </head>

  <body>
      <%

        dim c, pre, usu, nuevoPre

        pre = Request.QueryString("p")
        usu = Request.Cookies("Usuario")
        nuevoPre = Presupuesto_Secuencia(usu) 

        DuplicarEncabezado usu, pre, nuevoPre
        DuplicarDetalles usu, pre, nuevoPre

        response.redirect "../lista.asp"
    %>
  </body>
</html>

