<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim cc, sqlString, usuario, evento, calendario, contacto

    Usuario = Request.Cookies("Usuario")
    evento = Request.QueryString("ev")
    calendario = Request.QueryString("cal")
    contacto = Request.QueryString("con")

    if evento <> "*" then 
        sqlString = "DELETE FROM cal_Eventos_Participantes " & _
                     "WHERE (Usuario = '" & Usuario & "') " & _
                       "AND (evento = '" & evento & "') " & _
                       "AND (calendario = '" & calendario & "') " & _
                       "AND (contacto = '" & contacto & "');"

        set cc = Server.CreateObject("ADODB.Connection")
        cc.open Application("Conn")
            cc.execute(sqlString)
        cc.close: set cc = nothing
    end if

    response.redirect "cal_eventos_editar.asp?s=" & evento
%>