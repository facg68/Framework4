<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim c, t, sqlString

    sqlString = "DELETE FROM seg_Cripto_NumParse_Locales " & _
                "WHERE local = '" &  Request.QueryString("lm") & "';"

    set c = Server.CreateObject("ADODB.Connection")
    c.open Application("Conn")
        c.execute(sqlString)
    c.close: set c = nothing

    response.redirect "lista.asp"

response.write sqlString    
%>    