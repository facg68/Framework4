<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    response.Cookies("listaPreTipo") = request.form("cboTipoPresupuesto")
    response.Cookies("listaPreEstatus") = request.form("cboEstatusPresupuesto")
    response.Cookies("listaPreOrdenamiento") = request.form("ordenamiento")

    response.redirect "lista.asp"
%>