<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    '
    ' Duplicar la transacción seleccionada
    '
    dim con, sqlString, pre, vinculo

    pre = request.QueryString("p")
	 dia = Request.QueryString("d")

    sqlString = "INSERT INTO pre_Presupuesto_Detalles(Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, Contacto, Nota, NotaPre, NotaDonde, Incremento, Aplicado, Archivado) " & _
                "SELECT Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, Contacto, Nota, NotaPre, NotaDonde, Incremento, 0, 0 " & _
                  "FROM pre_Presupuesto_Detalles " & _
                 "WHERE Llave = " & Request.QueryString("l") & ";"    

    set con = Server.CreateObject("ADODB.Connection")
    con.open Application("Conn")
    con.execute(sqlString)
    con.close: set con = nothing

    vinculo = "pre_det_editar.asp?p=" & pre & "&d=" & request.QueryString("d") & "&v=" & request.QueryString("v") & "&t=" & request.QueryString("t") & "&e=" & request.QueryString("e") & "&o=" & request.QueryString("o")
    response.redirect vinculo
%> 