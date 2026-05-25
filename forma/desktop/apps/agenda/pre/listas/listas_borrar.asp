<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim c, Codigo

    Codigo = Request.QueryString("l")

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")
        cc.execute("DELETE FROM pre_Listas_Detalles WHERE Usuario = '" & Request.Cookies("Usuario") & "' AND Codigo = '" & Codigo & "';")
        cc.execute("DELETE FROM pre_Listas_Encabezado WHERE Usuario = '" & Request.Cookies("Usuario") & "' AND Codigo = '" & Codigo & "';")
    cc.close: set cc = nothing

    response.redirect "lista.asp"
%>