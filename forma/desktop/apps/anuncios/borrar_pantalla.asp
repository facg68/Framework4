<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim cc, tt, sqlString, sec

    sec = Request.QueryString("a")

    sqlString = "DELETE FROM seg_Anuncios_Pantallas " & _
                "WHERE (Pantalla = '" & sec & "');"

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")
        cc.execute(sqlString)
    cc.close: set cc = nothing

    response.redirect "pantallas.asp"
%> 