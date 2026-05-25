<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    dim cc, sqlString, usu, codigo, tipo

    usu = Request.Cookies("Usuario")
    tipo =  Request.QueryString("t")
    codigo = Request.QueryString("c")

    sqlString = "DELETE FROM con_Contactos_Categorias " & _
                "WHERE Usuario = '" & usu & "' " & _
                "AND Tipo = '" & tipo & "' " & _
                "AND Codigo = '" & Codigo & "';"

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")
        cc.execute(sqlString)
    cc.close: set cc = nothing

    response.redirect "cont_categorias.asp?t=" & tipo
%>