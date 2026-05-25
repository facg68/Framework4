<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim cc, tt, sqlString
    dim metadata, Paquete

    Paquete = Request.QueryString("p")
    Metadata = Request.QueryString("m")

    sqlString = "DELETE FROM discos_Paquetes_Metadata " &  _
                "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                "AND (Paquete = '" & Paquete & "') " & _
                "AND (Metadata = '" & MetaData & "');"

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")
        cc.execute(sqlString)
    cc.close: set cc = nothing

    response.redirect "editar.asp?m=" & paquete
%>