<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 

<%
    dim c, sqlString
    dim usuario, codigo

    usuario = Request.Cookies("usuario")
    codigo = Request.QueryString("con")

    sqlString = "UPDATE con_Contactos " & _
                "SET Estatus = 0 " & _
                "WHERE (Usuario = '" & usuario & "') " & _
                "AND (Codigo = '" & codigo & "');"

    set c = Server.CreateObject("ADODB.Connection")
    c.open Application("Conn")
        c.execute sqlString
    c.close: set c = nothing

    response.redirect "lista.asp"
%>
