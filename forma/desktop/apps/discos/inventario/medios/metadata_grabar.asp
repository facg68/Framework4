<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim cc, tt, sqlString
    dim metadata, Paquete

    Paquete = Request.QueryString("p")
    Metadata = Request.QueryString("m")

    sqlString = "INSERT INTO discos_Paquetes_Metadata(Usuario, Paquete, Metadata) " &  _
                "VALUES('" & Request.Cookies("Usuario") & "', '" & Paquete & "', '" & MetaData & "');"

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")
        cc.execute(sqlString)
    cc.close: set cc = nothing

    response.redirect "editar.asp?m=" & paquete
%>