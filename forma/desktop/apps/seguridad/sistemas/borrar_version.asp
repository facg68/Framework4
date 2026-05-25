<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim con, t, sqlString, Sistema, Version, OrdenadoPor

    Sistema = Request.QueryString("s")
    Version = Request.QueryString("v")
    OrdenadoPor = Request.QueryString("o")

    set con = Server.CreateObject("ADODB.Connection")
    con.open Application("Conn")

    sqlString = "DELETE FROM seg_Versiones " & _
                 "WHERE (Sistema = '" & Sistema & "') " & _
                   "AND (Version = '" & Version & "');"
    
    con.execute(sqlString)

    response.redirect "versiones.asp?s=" & Sistema & "&v=" & Version & "&0=" & OrdenadoPor
%>
