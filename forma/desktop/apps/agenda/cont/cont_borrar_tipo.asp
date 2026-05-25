<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    dim cc, sqlString
    dim codigo, usuario

    codigo = Request.QueryString("t")
    usuario = Request.QueryString("u")

    sqlString = "DELETE FROM con_Contactos_Tipos " & _
                "WHERE Usuario = '" & usuario & "' " & _
                "AND Codigo = '" & Codigo & "';"

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")
        cc.execute(sqlString)
    cc.close: set cc = nothing

    response.redirect "cont_tipos.asp"
%>