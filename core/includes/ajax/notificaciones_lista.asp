<%
    Response.ContentType = "application/json"
    On Error Resume Next  

    Dim con, t, sqlString, bloque, usuario

    usuario = Request.Cookies("Usuario")

    If Len(usuario) = 0 Then
        Response.Write "[]"
        Response.End
    End If

    sqlString = "SELECT Secuencia, Tipo, Titulo, Mensaje, Vinculo, Estatus " & _
                "FROM seg_Notificaciones " & _
                "WHERE Usuario = '" & usuario & "' " & _
                "ORDER BY Secuencia DESC;"   ' Más recientes primero

    Set con = Server.CreateObject("ADODB.Connection")
    con.Open Application("Conn")
        If Err.Number <> 0 Then
            Response.Write "[]"
            Response.End
        End If

        Set t = con.Execute(sqlString)
            If Err.Number <> 0 Then
                Response.Write "[]"
                Response.End
            End If

            bloque = "["

            Do While Not t.EOF
                bloque = bloque & "{""secuencia"":" & t("Secuencia") & _
                        ",""tipo"":""" & t("Tipo") & """" & _
                        ",""titulo"":""" & Replace(t("Titulo"), """", "\""") & """" & _
                        ",""mensaje"":""" & Replace(t("Mensaje"), """", "\""") & """" & _
                        ",""vinculo"":""" & Replace(t("Vinculo") & "", """", "\""") & """" & _
                        ",""estatus"":" & t("Estatus") & "},"
                t.MoveNext
            Loop

            If Right(bloque,1) = "," Then bloque = Left(bloque, Len(bloque)-1)
            bloque = bloque & "]"

        t.Close: Set t = Nothing
    con.Close: Set con = Nothing

    On Error GoTo 0 
    Response.Write bloque
%>