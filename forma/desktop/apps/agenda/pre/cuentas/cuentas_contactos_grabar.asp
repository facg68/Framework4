<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim c, t, sqlString
    dim Usuario, Cuenta, Contacto, Monto

    Usuario = Request.Cookies("Usuario")
    Cuenta = Request.Form("Cuenta")
    Contacto = Request.Form("contacto")
    Monto = Request.Form("Monto")

    set c = Server.CreateObject("ADODB.Connection")
    c.open Application("Conn")

    '
    ' Actualizamos las cantidades de los usuarios compartidos...
    '
    sqlString = "UPDATE pre_Cuentas_Comparticiones " & _
                   "SET MontoCompartido = " & Monto & _
                " WHERE (Usuario = '" & Usuario & "') " & _
                   "AND (Cuenta = '" & Cuenta & "');"

    c.execute(sqlString)

    '
    ' Añadimos el Contacto...
    '
    sqlString = "INSERT INTO pre_Cuentas_Comparticiones(Usuario, Cuenta, Contacto, MontoCompartido, UltimaFechaAplicada, Puntero) " & _
                "VALUES('" & Usuario & "', '" & Cuenta & "', '" & Contacto & "', " & Monto & ", NULL, 0);"

    c.execute(sqlString)

    '
    ' Cerramos la conexión y volvemos al editor...
    '
    c.close: set c = nothing

    response.redirect "cuentas_contactos.asp?c=" & cuenta
%>