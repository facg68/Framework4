<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 

<%
    dim c, sqlString
    dim usuario, codigo, ver, tipo, categ, orden1, orden2

    usuario = Request.Cookies("usuario")
    codigo = Request.QueryString("con")
    ver = Request.QueryString("v")
    tipo = Request.QueryString("t")
    categ = Request.QueryString("c")
    orden1 = Request.QueryString("o1")
    orden2 = Request.QueryString("o2")    

    sqlString = "UPDATE con_Contactos " & _
                "SET Estatus = 0 " & _
                "WHERE (Usuario = '" & usuario & "') " & _
                "AND (Codigo = '" & codigo & "');"

    set c = Server.CreateObject("ADODB.Connection")
    c.open Application("Conn")
        c.execute sqlString
    c.close: set c = nothing

    response.redirect "lista.asp?v=" & ver & "&t=" & tipo & "&c=" & categ & "&o1=" & orden1 & "&o2=" & orden2
%>
