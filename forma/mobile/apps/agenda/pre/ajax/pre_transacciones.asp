<%
    Response.ContentType = "application/json"
    On Error Resume Next  

    Dim con, t, sqlString, resultado, bloque

    usuario = Request.Cookies("Usuario")
    Presupuesto = Request.QueryString("Presupuesto")

    If Len(usuario) = 0 Then
        Response.Write "[]"
        Response.End
    End If

    sqlString = "SELECT ISNULL(COUNT(*), 0) AS Cuantos " & _
                  "FROM pre_Presupuesto_Detalles " & _
                 "WHERE (Usuario = '" & Usuario & "') " & _
                   "AND (Presupuesto = '" & Presupuesto & "') " & _
                   "AND (Aplicado = 1);"  

    Set con = Server.CreateObject("ADODB.Connection")
    con.Open Application("Conn")
        Set t = con.Execute(sqlString)
            if not (t.bof or t.eof) then
                bloque = "{""transacciones"":""" & t("Cuantos") & """}"
            end if
        t.Close: Set t = Nothing
    con.Close: Set con = Nothing

    On Error GoTo 0 
    Response.Write bloque
%>