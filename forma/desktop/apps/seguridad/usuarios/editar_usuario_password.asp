<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim cc, sqlString, usuario, ordenadoPor

    usuario = Request.Querystring("u")
    ordenadoPor = Request.Querystring("o")

    sqlString = "UPDATE seg_Usuarios SET usuReset = 1 WHERE usuCodigo = '" & usuario & "';"

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")
        cc.execute(sqlString)    
    cc.close: set cc = nothing

    Response.redirect "lista.asp?o=" & ordenadoPor    
%>