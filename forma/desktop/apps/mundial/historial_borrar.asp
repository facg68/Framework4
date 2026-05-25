<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim con, sqlString

    sqlString = "DELETE FROM mundial_Resultados WHERE Secuencia = " & request.QueryString("s") & ";"

    set con = Server.CreateObject("ADODB.Connection")
    con.Open Application("Conn")
        con.execute sqlString
    con.close: set con = nothing

    response.redirect "historial.asp"
%>