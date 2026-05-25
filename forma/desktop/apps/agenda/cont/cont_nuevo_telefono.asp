<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    dim c, sqlString
    dim usu, cont, tel, tipo

    usu = Request.Cookies("Usuario")
    cont = Request.QueryString("c")
    tel = Request.QueryString("t")
    tipo = Request.QueryString("l")

    tel = replace(tel, "*", "+")

    sqlString = "INSERT INTO con_Contactos_Telefonos(Usuario, Codigo, Telefono, Tipo) " & _
                "VALUES('" & usu & "', '" & cont & "', '" & tel & "', '" & tipo & "');"

    set c = Server.CreateObject("ADODB.Connection")
    c.open Application("Conn")
    c.execute sqlString
    c.close: set c = nothing

    response.redirect "cont_editar.asp?con=" & cont & "&tt=" & Request.QueryString("tt")
%>