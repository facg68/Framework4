<%
    dim con, t, usuario, snippet, estatus, sqlCommand

    usuario = Request.Cookies("Usuario")
    snippet = Request.QueryString("s")
    estatus = Request.QueryString("est")
    ventana = Request.QueryString("w")

    set con = Server.CreateObject("ADODB.Connection")
    con.open Application("Conn")

        if estatus <> "" then
            sqlCommand = "UPDATE seg_Usuarios_Snippets " & _
                            "SET snippetActivo = " & estatus & _
                        " WHERE (codUsuario = '" & usuario & "') " & _
                        " AND (snippet = '" & snippet & "');"

            con.execute(sqlCommand)
        else
            if ventana <> "" then
                sqlCommand = "UPDATE seg_Usuarios_Snippets " & _
                                "SET snippetMinimizado = " & ventana & _
                            " WHERE (codUsuario = '" & usuario & "') " & _
                            " AND (snippet = '" & snippet & "');"
                            
                con.execute(sqlCommand)        
            end if
        end if
       
    con.close: set con = nothing
%>