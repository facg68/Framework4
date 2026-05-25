<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    dim c, sqlString, sec, codigo, rel, contRel, cumple

    usu = Request.Cookies("Usuario")
    codigo = Request.QueryString("c")
    rel = Request.QueryString("r")
    contRel = Request.QueryString("k")
    cumple = Request.QueryString("q")

    sqlString = "INSERT INTO con_Contactos_Relacionados(Usuario, Codigo, Relacion, CodigoRelacionado, Cumple) " & _
                "VALUES('" & usu & "', '" & codigo & "', '" & rel & "', '" & contRel & "', '" & cumple & "');"

    set c = Server.CreateObject("ADODB.Connection")
    c.open Application("Conn")
    c.execute sqlString
    c.close: set c = nothing

    response.redirect "cont_editar.asp?con=" & codigo & "&tt=" & Request.QueryString("tt")
%>