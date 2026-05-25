<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    dim usuario, codigo, ver, tipo, categ, orden1, orden2
    dim c, sqlString

    usuario = Request.Cookies("usuario")
    codigo = Request.QueryString("con")
    ver = Request.QueryString("v")
    tipo = Request.QueryString("t")
    categ = Request.QueryString("c")
    orden1 = Request.QueryString("o1")
    orden2 = Request.QueryString("o2")

    sqlString = "UPDATE con_Contactos " & _
                   "SET Estatus = 1, " & _
                      " Visible = 1 " & _
                 "WHERE (Usuario = '" & usuario & "') " & _
                   "AND (Codigo = '" & codigo & "');"
    
    set c = Server.CreateObject("ADODB.Connection")
    c.open Application("Conn")
        c.execute(sqlString)
    c.close: set c = nothing    

    response.redirect "cont_editar.asp?con=" & Codigo & "&v=" & ver & "&t=" & tipo & "&c=" & categ & "&o1=" & orden1 & "&o2=" & orden2 
%>