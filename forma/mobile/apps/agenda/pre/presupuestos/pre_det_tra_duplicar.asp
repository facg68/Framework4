<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    '
    ' Duplicar la transacción seleccionada
    '
    dim con, sqlString, pre, vinculo

    registro = request.QueryString("registro")

    sqlString = "INSERT INTO pre_Presupuesto_Detalles(Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, Contacto, Nota, NotaPre, NotaDonde, Incremento, Aplicado, Archivado) " & _
                "SELECT Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, Contacto, Nota, NotaPre, NotaDonde, Incremento, 0, 0 " & _
                  "FROM pre_Presupuesto_Detalles " & _
                 "WHERE Llave = " & registro & ";"    

    set con = Server.CreateObject("ADODB.Connection")
    con.open Application("Conn")
    con.execute(sqlString)
    con.close: set con = nothing

    response.redirect "pre_det_editar.asp"
%> 