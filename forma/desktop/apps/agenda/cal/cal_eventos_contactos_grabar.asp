<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    dim cc, sqlString, usuario, evento, calendario, contacto

    Usuario = Request.Cookies("Usuario")
    evento = Request.QueryString("ev")
    calendario = Request.QueryString("cal")
    contacto = Request.QueryString("con")

    if evento <> "*" then 
        sqlString = "INSERT INTO cal_Eventos_Participantes(evento, usuario, calendario, contacto) " & _
                    "VALUES('" & evento & "', '" & Usuario & "', '" & calendario & "', '" & contacto & "');"

        set cc = Server.CreateObject("ADODB.Connection")
        cc.open Application("Conn")
            cc.execute(sqlString)
        cc.close: set cc = nothing
    end if

    response.redirect "cal_eventos_editar.asp?s=" & evento
%>