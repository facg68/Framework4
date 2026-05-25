<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    dim cc, tt, sqlString
    dim a, m, cal, v

    a = Request.QueryString("a")
    m = Request.QueryString("m")
    cal = Request.QueryString("c")
    v = Request.QueryString("v")

    if v = 0 then 
        v = 1
    else
        v = 0
    end if

    sqlString = "UPDATE cal_Calendarios " & _
                   "SET Seleccionado = " & v &  _
                " WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                   "AND (Codigo = '" & cal & "');"

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")
        cc.execute(sqlString)
    cc.close: set cc = nothing

    response.redirect "cal_calendario.asp?a=" & a & "&m=" & m
%>