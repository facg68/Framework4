<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    response.Cookies("lista_medios")("folder") =  Request.Form("cboFolder")
    response.Cookies("lista_medios")("tipo") = Request.Form("cbotipo")
    response.Cookies("lista_medios")("forma") = Request.Form("cboForma")
    response.Cookies("lista_medios")("plataforma") = Request.Form("cboPlataforma")
    response.Cookies("lista_medios")("amo") = Request.Form("txtAmo")
    response.Cookies("lista_medios")("ordenamiento") = Request.Form("cboOrden")

    Response.Cookies("lista_medios").Expires = Date() + 3000	

    response.Redirect "lista.asp"
%>