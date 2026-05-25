<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim cc, Comando, sqlString

    Usuario = Request.Cookies("Usuario")
    Comando = "exec dbo.pre_CerrarCuentasUsuario '" & Usuario & "';"

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")

        '
        ' Borramos las transacciones no aplicadas en presupuestos cerrados...
        '
        sqlString = "DELETE FROM dbo.pre_Presupuesto_Detalles " & _
                    "WHERE (Aplicado = 0) " & _
                    "AND (dbo.pre_EstatusPresupuesto(Usuario, Presupuesto) = 0)"

        cc.execute(sqlString)

        '
        ' Ahora realizamos el cierre de Cuentas del Usuario
        '
        cc.execute(Comando)

    cc.close: set cc = nothing

    response.redirect "lista.asp"
%>