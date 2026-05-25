<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    dim cc, sqlString
    dim codigo, Nombre, Usuario

    Usuario = Request.Cookies("Usuario")
    codigo = Request.QueryString("c")

    sqlString = "DELETE FROM cal_Calendarios " & _
                 "WHERE (Usuario = '" & Usuario & "') " & _
                 "AND (Codigo = '" & codigo & "');"

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")
        cc.execute(sqlString)
    cc.close: set cc = nothing

    response.redirect "cal_tipos.asp"
%>