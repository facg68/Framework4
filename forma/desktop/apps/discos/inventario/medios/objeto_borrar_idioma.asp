<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim cc, tt, sqlString, sec, ed, paquete, objeto

    sec = Request.QueryString("s")
    ed = Request.QueryString("ed")
    paquete = Request.QueryString("p")
    objeto = Request.QueryString("o")

    sqlString = "DELETE FROM discos_Objetos_Idiomas " & _
                "WHERE (Secuencia = '" & sec & "');"
    
    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")
        cc.execute(sqlString)
    cc.close: set cc = nothing

    response.redirect "editar_objeto.asp?p=" & Paquete & "&o=" & Objeto & "&e=" & ed
%>