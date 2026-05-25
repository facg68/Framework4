<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    '
    ' Init
    '
    dim c, t, sqlString, comando1, comando2, origen
 
    set c = Server.CreateObject("ADODB.Connection")
    c.open Application("Conn")

    comando1 = "UPDATE seg_Anuncios SET Orden = " & Request.QueryString("o2") & " WHERE Secuencia = " & Request.QueryString("s1") & ";"
    comando2 = "UPDATE seg_Anuncios SET Orden = " & Request.QueryString("o1") & " WHERE Secuencia = " & Request.QueryString("s2") & ";"
    origen = Request.QueryString("w")

    response.write comando1 & "<br/><br/>"
    response.write comando2 & "<br/><br/>"

    '
    ' Ejecutamos los comandos...
    '

    c.execute(comando1)
    c.execute(comando2)

    c.close: set c = nothing

    if origen = 1 then
        response.redirect "lista_total.asp?tv=" & request.QueryString("tv") & "&e=" & request.QueryString("e") & "&op=" & request.QueryString("op")
    else
        response.redirect "lista.asp?tv=" & request.QueryString("tv") & "&e=" & request.QueryString("e") & "&op=" & request.QueryString("op")
    end if
%>