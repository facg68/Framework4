<!-- #include virtual="/core/includes/kernel/local.inc" -->

<html>
    <head>
        <%
            Function BytesToString(bytes)
                Dim stream
                Set stream = Server.CreateObject("ADODB.Stream")

                stream.Type = 1
                stream.Open
                stream.Write bytes
                stream.Position = 0

                stream.Type = 2
                stream.Charset = "utf-8"

                BytesToString = stream.ReadText

                stream.Close
                Set stream = Nothing
            End Function        
        %>
    </head>

    <body>
        <%
            Dim Conn, Usuario, rawData, registros, r, campos
            Dim snippetNombre, snippetTop, snippetLeft, snippetIndex
            Dim sql

            Usuario = Request.Cookies("Usuario")

            ' Leer los datos enviados por sendBeacon
            rawData = Request.BinaryRead(Request.TotalBytes)
            rawData = BytesToString(rawData)

            Set Conn = Server.CreateObject("ADODB.Connection")
            Conn.Open Application("Conn")

            ' Separar cada ventana
            registros = Split(rawData, ";")

            For Each r In registros
                If Len(Trim(r)) > 0 Then
                    campos = Split(r, "|")

                    snippetNombre = campos(0)
                    snippetTop    = campos(1)
                    snippetLeft   = campos(2)
                    snippetIndex  = campos(3)

                    sql = "UPDATE seg_Usuarios_Snippets " & _
                            "SET snippetTop=" & snippetTop & ", " & _
                                " snippetLeft=" & snippetLeft & ", " & _
                                " snippetIndex=" & snippetIndex & " " & _
                        "WHERE codUsuario='" & Usuario & "' " & _
                            "AND snippet='" & snippetNombre & "'"

                    Conn.Execute sql
                End If
            Next

            Conn.Close: Set Conn = Nothing
        %>    
    </body>
</html>
