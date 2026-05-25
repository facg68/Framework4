<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim cc, sqlString, secuencia, Usuario

    Usuario = Request.Cookies("Usuario")
    secuencia = Request.QueryString("s")

    sqlString = "DELETE FROM cal_Eventos " & _
                 "WHERE (Usuario = '" & Usuario & "') " & _
                 "AND (Secuencia = " & Secuencia & ");"

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")
        cc.execute(sqlString)
    cc.close: set cc = nothing

    response.redirect "cal_calendario.asp"
%>