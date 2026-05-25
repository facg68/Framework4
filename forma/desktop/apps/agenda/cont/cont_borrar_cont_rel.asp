<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    dim c, sqlString, sec, cont

    sec = Request.QueryString("s")
    cont = Request.QueryString("c")

    sqlString = "DELETE FROM con_Contactos_Relacionados " & _
                 "WHERE (Secuencia = " & sec & ");"

    set c = Server.CreateObject("ADODB.Connection")
    c.open Application("Conn")
    c.execute sqlString
    c.close: set c = nothing

    response.redirect "cont_editar.asp?con=" & cont & "&tt=" & Request.QueryString("tt")
%>