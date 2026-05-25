<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function LimpiarApostrofes(valor)
                LimpiarApostrofes = Replace(valor,"'","´")
                LimpiarApostrofes = Replace(valor,"yy","&")                
            end function
        %>
    </head>

    <body>
        <%
            dim cc, tt, sqlString, usuario, paquete, objeto, editor
            dim Opcion, titulo, exito, lado

            Usuario = Request.Cookies("Usuario")
            Opcion = Request.QueryString("op")
            Paquete = Request.QueryString("p")
            Objeto = Request.QueryString("o")
            Editor = Request.QueryString("ed")

            Select Case Opcion
                Case 1
                    '
                    ' Objeto Musical de un solo lado
                    '
                    titulo = LimpiarApostrofes(Request.QueryString("t"))
                    exito = Request.QueryString("e")
                    if exito = "" then exito = 0

                    sqlString = "INSERT INTO discos_Objetos_Detalle(Usuario, Paquete, Objeto, Titulo, NumSerieLLave, Exito, Lado) " & _
                                "VALUES('" & Usuario & "', '" & Paquete & "', '" & Objeto & "', '" & Titulo & "', NULL, " & Exito & ", NULL);"

                Case 2
                    '
                    ' Objeto Musical Multilado
                    '
                    lado =  Request.QueryString("la")
                    titulo = LimpiarApostrofes(Request.QueryString("t"))
                    exito = Request.QueryString("e")
                    if exito = "" then exito = 0

                    sqlString = "INSERT INTO discos_Objetos_Detalle(Usuario, Paquete, Objeto, Titulo, NumSerieLLave, Exito, Lado) " & _
                                "VALUES('" & Usuario & "', '" & Paquete & "', '" & Objeto & "', '" & Titulo & "', NULL, " & Exito & ", '" & lado & "');"
                
                Case 3
                    '
                    ' Libro - Capitulos
                    '
                    titulo = LimpiarApostrofes(Request.QueryString("t"))

                    sqlString = "INSERT INTO discos_Objetos_Detalle(Usuario, Paquete, Objeto, Titulo, NumSerieLLave, Exito, Lado) " & _
                                "VALUES('" & Usuario & "', '" & Paquete & "', '" & Objeto & "', '" & Titulo & "', NULL, 0, NULL);"

                Case 4
                    '
                    ' Juegos, Software - Numeros de Serie
                    '
                    titulo = LimpiarApostrofes(Request.QueryString("t"))
                    num = Request.QueryString("num")

                    sqlString = "INSERT INTO discos_Objetos_Detalle(Usuario, Paquete, Objeto, Titulo, NumSerieLLave, Exito, Lado) " & _
                                "VALUES('" & Usuario & "', '" & Paquete & "', '" & Objeto & "', '" & Titulo & "', '" & num & "', 0, NULL);"

                Case 5
                    '
                    ' Peliculas - Capitulos
                    '
                    titulo = LimpiarApostrofes(Request.QueryString("t"))

                    sqlString = "INSERT INTO discos_Objetos_Detalle(Usuario, Paquete, Objeto, Titulo, NumSerieLLave, Exito, Lado) " & _
                                "VALUES('" & Usuario & "', '" & Paquete & "', '" & Objeto & "', '" & Titulo & "', NULL, 0, NULL);"

                Case 6
                    '
                    ' Peliculas - Protagonistas
                    '
                    Protagonista = LimpiarApostrofes(Request.QueryString("prot"))

                    sqlString = "INSERT INTO discos_Objetos_Protagonistas(Usuario, Paquete, Objeto, Protagonista) " & _
                                "VALUES('" & Usuario & "', '" & Paquete & "', '" & Objeto & "', '" & Protagonista & "');"

                Case 7
                    '
                    ' Peliculas - Idiomas
                    '
                    Idioma = Request.QueryString("idi")
                    Audio = Request.QueryString("aud")
                    Subtitulo = Request.QueryString("sub")

                    sqlString = "INSERT INTO discos_Objetos_Idiomas(Usuario, Paquete, Objeto, Idioma, Audio, Subtitulos) " & _
                                "VALUES('" & Usuario & "', '" & Paquete & "', '" & Objeto & "', '" & Idioma & "', " & Audio & ", " & Subtitulo & ");"

            End Select

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")  
                cc.execute(sqlString)
            cc.close: set cc = nothing

            response.redirect "editar_objeto.asp?p=" & Paquete & "&o=" & Objeto & "&e=" & Editar
        %>
    </body>
</html>