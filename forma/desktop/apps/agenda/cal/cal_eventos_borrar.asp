<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    dim cc, sqlString, dia, mes, amo
    dim secuencia, Usuario, vinculo

    Usuario = Request.Cookies("Usuario")
    secuencia = Request.QueryString("s")
    dia = Request.QueryString("d")
    mes = Request.QueryString("m")
    amo = Request.QueryString("a")

    vinculo = "cal_eventos.asp?d=" & dia & "&m=" & mes & "&a=" & amo

    sqlString = "DELETE FROM cal_Eventos " & _
                 "WHERE (Usuario = '" & Usuario & "') " & _
                 "AND (Secuencia = " & Secuencia & ");"

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")
    cc.execute(sqlString)
    cc.close: set cc = nothing

    response.redirect vinculo
%>