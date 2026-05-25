<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    dim c, sqlString, sec, cont

    cont = Request.QueryString("c")
    sec = Request.QueryString("s")

    sqlString = "DELETE FROM con_Contactos_Telefonos " & _
                 "WHERE Secuencia = " & sec & ";"

    set c = Server.CreateObject("ADODB.Connection")
    c.open Application("Conn")
    c.execute sqlString
    c.close: set c = nothing

    response.redirect "cont_editar.asp?con=" & cont & "&tt=" & Request.QueryString("tt")
%>