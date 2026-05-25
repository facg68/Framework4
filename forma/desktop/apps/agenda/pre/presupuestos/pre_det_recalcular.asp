<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    '
    ' Recalcular todas las transacciones
    '
    dim con, t, pr, sqlString, pre, usu, mOrigen, mDestino

    pre = Request.QueryString("p")
    usu = Request.Cookies("Usuario")

    set con = Server.CreateObject("ADODB.Connection")
    con.open Application("Conn")

    set pr = con.execute("SELECT MonedaOrigen, MonedaDestino FROM pre_Presupuesto_Encabezado " & _
                          "WHERE Usuario = '" & usu & "' " & _
                            "AND Presupuesto = '" & pre & "';") 

    if not (pr.bof or pr.eof) then
        mOrigen = pr("MonedaOrigen")
        mDestino = pr("MonedaDestino")

        set t = con.execute("SELECT * FROM pre_Presupuesto_Detalles " & _
                             "WHERE Usuario = '" & usu & "' " & _
                               "AND Presupuesto = '" & pre  & "';")

        if not (t.bof or t.eof) then
            Do
                sqlString = "UPDATE pre_Presupuesto_Detalles " & _
                               "SET MontoDestino = (-1 * MontoOrigen)," & _
                                  " MontoCambio =  dbo.Cripto_CambiarMoneda((-1 * MontoOrigen),'" & mOrigen & "','" & mDestino & "') " & _
                             "WHERE Presupuesto = '" & pre & "' " & _
                               "AND Usuario = '" & usu & "';"

                con.execute(sqlString)
                t.MoveNext
            Loop Until t.eof
        end if

        t.close: set t = nothing
    end if

    pr.close: set pr = nothing
    con.close: set con = nothing

    response.redirect "pre_det_editar.asp?p=" & pre
%> 