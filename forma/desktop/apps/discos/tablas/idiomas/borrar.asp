<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim cc, tt, sqlString, Codigo

    Codigo = Request.QueryString("c")

    sqlString = "DELETE FROM discos_Idiomas " & _
                "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                "AND (Codigo = '" & Codigo & "');"

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")
        cc.execute(sqlString)
    cc.close: set cc = nothing

    response.redirect "lista.asp"
'response.write sqlString    
%>    