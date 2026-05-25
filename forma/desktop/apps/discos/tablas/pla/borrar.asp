<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim cc, tt, sqlString, Usuario, Codigo, Estatus, Tipo, Ordenamiento

    Usuario = Request.Cookies("Usuario")
    Codigo = Request.QueryString("c")
    Estatus = Request.QueryString("e")
    Tipo = Request.QueryString("t")
    Ordenamiento = Request.QueryString("o")

    sqlString = "DELETE FROM discos_Plataformas " & _
                "WHERE (Usuario = '" & Usuario & "') " & _
                "AND (Codigo = '" & Codigo & "');"

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")
        cc.execute(sqlString)
    cc.close: set cc = nothing

    response.redirect "lista.asp?e=" & Estatus & "&t=" & Tipo & "&o=" & Ordenamiento
%>    