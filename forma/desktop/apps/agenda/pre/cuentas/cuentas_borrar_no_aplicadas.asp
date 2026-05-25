<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim cc, sqlString

    Usuario = Request.QueryString("Usuario")

    sqlString = "DELETE FROM dbo.pre_Presupuesto_Detalles " & _
                "WHERE (Aplicado = 0) " & _
                "AND (dbo.pre_EstatusPresupuesto(Usuario, Presupuesto) = 0)"

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")
      cc.execute(sqlString)
    cc.close: set cc = nothing

    response.redirect "lista.asp"
%>