<%
    dim con, secuencia

    secuencia = Request("s")

    Set con = Server.CreateObject("ADODB.Connection")
    con.Open Application("Conn")    
        If IsNumeric(secuencia) Then
            con.Execute "UPDATE seg_Notificaciones SET Estatus = 0 WHERE secuencia = " & secuencia
        End If
    con.close: set con = nothing
%>