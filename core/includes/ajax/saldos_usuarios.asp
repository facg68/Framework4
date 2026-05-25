<%
    Response.ContentType = "application/json"
    On Error Resume Next  

    Dim conn, t, s, sqlString, resultado, bloque, usuario

    usuario = Request.Cookies("Usuario")

    If Len(usuario) = 0 Then
        Response.Write "[]"
        Response.End
    End If

    set conn = Server.CreateObject("ADODB.Connection")
    conn.Open Application("Conn") 

        If Err.Number <> 0 Then
            Response.Write "[]"
            Response.End
        End If
    
        '
        ' Puede Ver Saldos?
        '

        set t = conn.execute("SELECT usuVerSaldos FROM seg_Usuarios WHERE usuCodigo = '" & usuario & "';")
            If Err.Number <> 0 Then
                Response.Write "[]"
                Response.End
            End If

            if not (t.bof or t.eof) then
                if t("usuVerSaldos") = 1 then
                    '
                    ' Puede Ver Los Saldos... Preparamos la Data
                    '
                    sqlString = "SELECT Nombre, dbo.pre_SaldoCuenta(Usuario, Codigo, LocalMonetario) AS Monto " & _
                                "FROM pre_Cuentas " & _
                                "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                "AND (TipoCuenta = 'A') " & _
                                "AND (Grupo = 'A') " & _
                                "AND ((Codigo = 'PRE-000') OR (Codigo = 'EF-000')) " & _
                                "ORDER BY Nombre;"  

                    set s = conn.execute(sqlString)
                        If Err.Number <> 0 Then
                            Response.Write "[]"
                            Response.End
                        End If

                        if not (s.bof or s.eof) then
                            bloque = "["

                            Do While Not s.EOF
                                bloque = bloque & "{""nombre"":""" & s("Nombre") & """" & _
                                                  ",""monto"":""" & FormatNumber(s("Monto")) & """},"
                                s.MoveNext
                            Loop

                            If Right(bloque,1) = "," Then bloque = Left(bloque, Len(bloque)-1)
                            bloque = bloque & "]"
                        end if
                    s.close: set s = nothing
                end if
            end if
        t.close: set t = nothing
    conn.close: set conn = nothing

    On Error GoTo 0 
    Response.Write bloque
%>        