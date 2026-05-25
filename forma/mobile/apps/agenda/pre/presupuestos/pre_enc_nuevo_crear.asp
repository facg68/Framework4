<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            Function NuevoCodigo(Usuario)
                dim c, tt, sqlString, max

                sqlString = "SELECT ISNULL(MAX(CAST(RIGHT(Presupuesto, 9) AS numeric(9, 0))), 0) AS Maximo " & _
                            "FROM pre_Presupuesto_Encabezado " & _
                            "WHERE (LEFT(Presupuesto, 3) = 'PR-') " & _
                            "AND Usuario = '" & Usuario & "';"
                                 
                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                set tt = c.execute(sqlString)

                if (tt.bof or tt.eof) then
                    NuevoCodigo = "PR-000000001"
                else
                    max = cint(tt("Maximo")) + 1 
                    NuevoCodigo = "PR-" & RIGHT("000000000" + cstr(max), 9)
                end if

                tt.close: set tt = nothing
                c.close: set c = nothing
            End Function

            Function TipoPlantilla(Usuario, Plantilla)
                dim c, tt, sqlString

                sqlString = "SELECT Tipo " & _
                            "FROM pre_Presupuesto_Encabezado " & _
                            "WHERE Presupuesto = '" & Plantilla & "' " & _
                            "AND Usuario = '" & Usuario & "';"
                                 
                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                set tt = c.execute(sqlString)

                if (tt.bof or tt.eof) then
                    TipoPlantilla = "P"
                else
                    TipoPlantilla =  tt("Tipo")
                end if

                tt.close: set tt = nothing
                c.close: set c = nothing
            End Function            

            function FechaServer(fechaFormulario)
                dim d, m, a

                d = left(fechaFormulario, 2)
                m = mid(fechaFormulario, 4, 2)
                a = right(fechaFormulario, 4)

                FechaServer = a & "-" & right("0" & m, 2) & "-" & right("0" & d, 2)
            end function     

            function DateFormat(FechaVBS)
                dim p1, p2, p3, k, p, cadena, segmento, inicio

                '
                ' Parsear a mano... Uff!!!
                '

                p = 0
                cadena = FechaVBS & "/"
                inicio = 1

                for k = 1 to len(cadena)
                    if mid(cadena, k, 1) = "/" then
                        p = p + 1
                        segmento = mid(cadena, inicio, (k - inicio))
                        inicio = k + 1

                        select case p
                            case 1: p1 = right("0" & segmento, 2)
                            case 2: p2 = right("0" & segmento, 2)
                            case 3: p3 = segmento
                        end select
                    end if
                next

                DateFormat = p3 & "-" & p1 & "-" & p2
            end function                  

            Function SaldoAnterior(Usuario, Presupuesto)   
                dim c, tt, sqlString, max

                sqlString = "SELECT dbo.pre_Enc_SaldoFinal('" & Usuario & "', '" & Presupuesto & "') AS Saldo;"
                                 
                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                set tt = c.execute(sqlString)

                if (tt.bof or tt.eof) then
                    SaldoAnterior =  0
                else
                    SaldoAnterior = tt("Saldo")
                end if

                tt.close: set tt = nothing
                c.close: set c = nothing            
            End Function

            Function EfectivoAnterior(Usuario, Presupuesto)   
                dim c, tt, sqlString, max

                sqlString = "SELECT dbo.pre_Enc_EfectivoFinal('" & Usuario & "', '" & Presupuesto & "') AS Saldo;"
                                 
                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                set tt = c.execute(sqlString)

                if (tt.bof or tt.eof) then
                    EfectivoAnterior =  0
                else
                    EfectivoAnterior = tt("Saldo")
                end if

                tt.close: set tt = nothing
                c.close: set c = nothing            
            End Function            

            Function NombreContacto(Usuario, Contacto)   
                dim c, tt, sqlString, max

                sqlString = "SELECT PrimerNombre + ' ' + PrimerApellido AS NombreContacto " & _
                              "FROM con_Contactos " & _
                             "WHERE (Usuario = '" & Usuario & "') " & _
                               "AND (Codigo = '" & Contacto & "');"
                                 
                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                set tt = c.execute(sqlString)

                if (tt.bof or tt.eof) then
                    NombreContacto =  " - - ERROR - - "
                else
                    NombreContacto = tt("NombreContacto")
                end if

                tt.close: set tt = nothing
                c.close: set c = nothing            
            End Function            

            Function LabelCuentasCompartidas(Usuario, Cuenta, NombreCuenta, Clase)
                dim c, tt, sqlString

                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")

                if Clase <> "N" then
                    set tt = c.execute("SELECT Secuencia, Contacto, MontoCompartido, UltimaFechaAplicada, Puntero " & _
                                         "FROM pre_Cuentas_Comparticiones " & _
                                   "WHERE Usuario = '" & Usuario & "' " & _
                                     "AND Cuenta = '" & Cuenta & "' " & _
                                "ORDER BY Secuencia ASC;")

                    if (tt.bof or tt.eof) then
                        LabelCuentasCompartidas = NombreCuenta
                    else
                        tt.MoveFirst

                        Monto = tt("MontoCompartido")
                        Puntero = 0
                        NuevoPuntero = 0

                        '
                        ' Parte 1 - Donde está el puntero?
                        '

                        Do
                            If (Puntero = 0) And (cint(tt("Puntero")) = 1) Then
                                Puntero = cint(tt("Secuencia"))
                            End If

                            tt.MoveNext
                        Loop Until tt.EOF

                        If Puntero > 0 Then
                            '
                            ' Parte 2 - Hay algún registro luego del puntero?
                            '

                            tt.MoveFirst

                            Do
                                If ( cint(tt("Secuencia")) > cint(Puntero) ) Then
                                    If NuevoPuntero = 0 Then
                                        NuevoPuntero = cint(tt("Secuencia"))
                                    End If
                                End If

                                tt.MoveNext
                            Loop Until tt.EOF

                            '
                            ' Parte 3 - Si el NuevoPuntero es igual a 0, entonces el NuevoPuntero
                            '           es el primer registo de nuestra lista
                            '
                            tt.MoveFirst

                            If NuevoPuntero = 0 Then
                                NuevoPuntero = cint(tt("Secuencia"))
                            End If

                            '
                            ' Parte 4 - Ahora que tenemos el puntero, procedemos
                            '           a crear el nuevo label
                            '
        
                            Do
                                If cint(tt("Secuencia")) = cint(NuevoPuntero) Then
                                    LabelCuentasCompartidas = "A " & NombreContacto(usuario, tt("Contacto")) & " le toca pagar " & Monto
                                End If

                                tt.MoveNext
                            Loop Until tt.EOF

                        End If

                        tt.close: set tt = nothing
                    End If
                Else
                    set tt = c.execute("SELECT Nombre " & _
                                         "FROM pre_Cuentas " & _
                                        "WHERE Usuario = '" & Usuario & "' " & _
                                          "AND Codigo = '" & Cuenta & "';")
                    
                    If Not (tt.BOF Or tt.EOF) Then
                        LabelCuentasCompartidas = tt("Nombre")
                    End If
                end if

                c.close: set c = nothing
            end function

            Function TipoCuentasCompartida(Usuario, Cuenta)
                dim c, tt, sqlString

                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")

                set tt = c.execute("SELECT Clase " & _
                                     "FROM pre_Cuentas " & _
                                    "WHERE (Usuario = '" & Usuario & "') " & _
                                      "AND (Codigo = '" & Cuenta & "');")

                if (tt.bof or tt.eof) then
                    TipoCuentasCompartida = "N"
                else
                    TipoCuentasCompartida = tt("Clase")
                end if

                tt.close: set tt = nothing
                c.close: set c = nothing
            end function   

            Function FechaHoy()
                dim a, m, d

                d = RIGHT("00" & day(DATE()) ,2)
                m = RIGHT("00" & month(DATE()), 2)
                a = year(DATE())

                FechaHoy = a & "-" & m & "-" & d
            end function   

            Function HoraHoy()
                dim h, m
                
                h = RIGHT("00" & Hour(Time()), 2)
                m = RIGHT("00" & Minute(Time()), 2)

                HoraHoy = h & m
            end function 

            function fLoop(Fecha)
                dim d, m, a

                d = day(Fecha)
                m = month(Fecha)
                a = year(Fecha)

                fLoop = a & "-" & right("00" & m, 2) & "-" & right("00" & d, 2) 
            end function            

            sub EventosCalendario(Presupuesto, FechaDesde, FechaHasta)      
                dim Fecha, FechaFinal, sw
                dim cc, tt, sqlString, cmdString

                Fecha = fLoop(FechaDesde)
                FechaFinal = fLoop(FechaHasta)
                sw = 0

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")

                    Do
                        sqlString = "exec pre_Nuevo_Eventos '" & request.Cookies("Usuario") & "', '" & Fecha & "'"
                        set tt = cc.execute(sqlString)

                        if not (tt.bof or tt.eof) then
                            Do
                                if tt("DbCr") = 0 then
                                    cOrigen = "PRE-000"
                                    cDestino = "SYS-000"
                                else
                                    cOrigen = "SYS-000"
                                    cDestino = "PRE-000"
                                end if

                                cmdString = "INSERT INTO pre_Presupuesto_Detalles(Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, Aplicado, Incremento, Archivado, Contacto) " & _
                                            "VALUES('" & Presupuesto & "', '" & Request.Cookies("Usuario") & "','" & Fecha & "'," & tt("Hora") & ", '" & cOrigen & "', " & (-1 * tt("Monto")) & "," & _
                                                "'" & tt("Descripcion") & "', '" & cDestino & "', " & tt("Monto") & ", " & tt("Monto") & ", 0, 0, 0, "

                                if PrimerContactoEvento(tt("Llave")) = "" then
                                    cmdString = cmdString & "NULL"
                                else
                                    cmdString = cmdString & "'" & PrimerContactoEvento(tt("Llave")) & "'"
                                end if

                                cmdString = cmdString & "); "

                                cc.execute(cmdString)                                            
                                tt.MoveNext
                            Loop Until tt.eof
                        end if

                        tt.close: set tt = nothing

                        Fecha = fLoop(DateAdd("d", 1, Fecha))
                        if (Fecha = FechaFinal) then sw = 1
                    Loop Until (sw = 1)

                cc.close: set cc = nothing                
            end sub

            Function PrimerContactoEvento(Secuencia)
                dim cc, tt, sqlString

                sqlString = "SELECT Top 1 Contacto " & _
                              "FROM cal_Eventos_Participantes " & _
                             "WHERE Evento = " & Secuencia & ";"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.execute(sqlString)

                    if not (tt.bof or tt.eof) then
                        PrimerContactoEvento = tt("Contacto")
                    else
                        PrimerContactoEvento = NULL
                    end if

                tt.close: set tt = nothing
                cc.close: set cc = nothing                
            End Function
        %>
    </head>

    <body>
        <%
            dim con, t, sqlString, usu, pre
            dim desde, hasta, multi, tmodelo

            dim cMes, cAmo, cPlantilla, cDesde, cHasta
            dim cOrigen, cDestino, cNombre, cAnterior, cTipo
            dim cReglas, incActual, tpActual, ff, dIncremento

            '-----------------------------------------
            ' Cargamos los datos desde el formulario '
            '-----------------------------------------

            cMes = Request.form("mes")
            cAmo = Request.form("amo")
            cPlantilla = Request.form("Plantilla")
            cDesde = Request.form("desde")
            cHasta = Request.form("hasta")
            cNombre = Request.form("nuevoPre")
            cOrigen = Request.form("monedaOrigen")
            cDestino = Request.form("monedaDestino")
            cAnterior = Request.form("preAnterior")
            cReglas = Request.form("reglas")  

            '----------------------------------------
            ' Verificamos y corregimos los valores  '
            '----------------------------------------            

            usu = Request.Cookies("Usuario")
            pre = NuevoCodigo(usu)
            desde = FechaServer(cDesde)
            hasta = FechaServer(cHasta)
            cTipo = TipoPlantilla(usu, cPlantilla)            

            if (cOrigen = cDestino) then
                multi = 0
            else
                multi = 1
            end if

            '-----------------------------------------
            ' Abrimos la conexion con el servidor... '
            '-----------------------------------------            

            set con = server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            '-----------------------------------------------------
            ' Procesamos y vamos creando el nuevo presupuesto... '
            '-----------------------------------------------------

            if cPlantilla = "*" then

                '------------------------------------'
                ' Se va a crear un presupuesto vacío '
                '------------------------------------'

                sqlString = "INSERT INTO pre_Presupuesto_Encabezado(Presupuesto, Usuario, Nombre, Tipo, Desde, Hasta, SaldoFinal, MultiPrecio, MonedaOrigen, MonedaDestino, Estatus, Cuantificable, Obsoleto) " & _
                            "VALUES ('" & pre & "', '" & usu & "','" & cNombre & "', 'P', '" & desde & "', '" & hasta & "', 0.00, " & multi & ", '" & cOrigen & "','" & cDestino & "', 1, 1, 0);"
                
                con.Execute (sqlString)

            else
                '----------------------------------------------------------------------------'
                ' Los Modelos se procesan insertando los registros de uno en uno para poder  '
                ' cambiar las fechas de los eventos según el "incremento" especificado       '
                ' en el detalle del modelo seleccionado                                      '
                '                                                                            '
                ' PARTE 1: Encabezado de Presupuesto                                         '
                '----------------------------------------------------------------------------'                

                sqlString = "INSERT INTO pre_Presupuesto_Encabezado(Presupuesto, Usuario, Nombre, Tipo, Desde, Hasta, SaldoFinal, MultiPrecio, MonedaOrigen, MonedaDestino, Estatus, Cuantificable, Obsoleto) " & _
                                    "VALUES ('" & pre & "', '" & usu & "','" & cNombre & "', 'P', '" & desde & "', '" & hasta & "', 0.00, " & multi & ", '" & cOrigen & "','" & cDestino & "', 1, 1, 0);"                    
                con.Execute sqlString

                '----------------------------------------------------------------------------'
                ' PARTE 2: Detalle de Presupuesto                                            '
                '----------------------------------------------------------------------------'
                
                sqlString = "SELECT '" & pre & "' AS Presupuesto, '" & usu & "' AS Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, Contacto, Nota, NotaPre, NotaDonde, Incremento " & _
                            "FROM pre_Presupuesto_Detalles " & _
                            "WHERE (Presupuesto = '" & cPlantilla & "') AND (Usuario='" & usu & "') " & _
                        "ORDER BY Fecha, Hora ASC;"
                
                Set tmodelo = con.execute(sqlString)                                               

                fModelo = mid(desde, 6,2) & "/" & right(desde, 2) & "/" & left(desde, 4)

                If Not (tmodelo.BOF Or tmodelo.EOF) Then
                    Do
                        fModelo = DateAdd("d", tmodelo("Incremento"), fModelo)
                        tFecha = DateFormat(fModelo)
                        
                        sqlString = "INSERT INTO pre_Presupuesto_Detalles(Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, Contacto," & _
                                                                        " Nota, NotaPre, NotaDonde, Aplicado, Incremento, Archivado) " & _
                                    "VALUES('" & tmodelo("Presupuesto") & "', '" & tModelo("Usuario") & "','" & tFecha & "'," & tModelo("Hora") & "," & _
                                    "'" & tModelo("CuentaOrigen") & "', " & tModelo("MontoOrigen") & ",'" & tModelo("Descripcion") & "'," & _
                                    "'" & tModelo("CuentaDestino") & "', " & tModelo("MontoDestino") & ", " & tModelo("MontoCambio") & ", " & _
                                    "'" & tModelo("Contacto") & "', '" & tModelo("Nota") & "', " & tModelo("NotaPre") & ", '" & tModelo("NotaDonde") & "', 0, 0, 0);"

                        con.Execute sqlString

                        tmodelo.MoveNext
                    Loop Until tmodelo.EOF
                End If

                tmodelo.close: set tModelo = nothing

                '-----------------------------------------------'
                '                                               '
                ' Verificamos si se deben aplicar las reglas de '
                ' creación de presupuestos                      '
                '                                               '
                '-----------------------------------------------'       

                if cReglas = 1 then
                    '-----------------------------------------------'
                    '                                               '
                    ' PARTE 3: Las cuentas anuales que apliquen...  '
                    '                                               '
                    '-----------------------------------------------'
                    sqlString = "exec pre_TransaccionesAnuales '" & usu & "', " & cAmo & ", '" & desde & "','" & Hasta & "'"

                    set tmodelo = con.Execute(sqlString)

                    If Not (tmodelo.BOF Or tmodelo.EOF) Then
                        Do
                            label = LabelCuentasCompartidas(usu, tmodelo("Cuenta"), tmodelo("NombreCuenta"), tmodelo("Clase"))                        
                            mMonto = tmodelo("Monto") 
                            if mMonto < 0 then mMonto = (-1 * mMonto)

                            if TipoCuentasCompartida(usu, tmodelo("Cuenta")) = "R" then
                                if InStr(Label, NombreContacto(usu, usu)) = 0 then
                                    '---------------------------------------------------'
                                    '                                                   '
                                    ' Al usuario NO LE TOCA PAGAR, por lo que el monto  '
                                    ' es 0 (a otra persona le toca hacer el pago)       '
                                    '                                                   '
                                    '---------------------------------------------------'
                                    mMonto = 0
                                end if
                            end if

                            sqlString = "INSERT INTO pre_Presupuesto_Detalles(Presupuesto, Usuario, Fecha, Hora, CuentaOrigen," & _
                                                   " MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, Contacto, Aplicado, Incremento, Archivado) " & _
                                        "VALUES('" & pre & "', '" & usu & "','" & tmodelo("Fecha") & "', 700, 'PRE-000', " & (-1 * mMonto) & ", '" & Label & "', " & _
                                               "'" & tmodelo("Cuenta") & "', " & mMonto & ", " & mMonto & ", '" & tmodelo("Contacto") & "', 0, 0, 0);"

                            con.Execute sqlString

                            tmodelo.MoveNext
                        Loop Until tmodelo.EOF
                    end If                            

                    '---------------------------------------------------------------'
                    '                                                               '
                    ' PARTE 4: Los productos presupuestados para el periodo actual  '
                    '                                                               '
                    '---------------------------------------------------------------'                    

                    sqlString = "INSERT INTO pre_Presupuesto_Detalles(Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, ItemLista, Contacto, Aplicado, Incremento, Archivado) " & _
                                "SELECT '" & pre & "' AS Presupuesto, '" & usu & "' AS Usuario, Fecha, 700 AS Hora,'PRE-000' AS CuentaOrigen, - dbo.Cripto_CambiarMoneda(Precio, LocalPrecio, '" & cOrigen & "') AS MontoOrigen, Nombre AS Descripcion, 'PRE-000' AS CuentaDestino, " & _
                                "dbo.Cripto_CambiarMoneda(Precio, LocalPrecio, '" & cOrigen & "') AS MontoDestino, dbo.Cripto_CambiarMoneda(Precio, LocalPrecio, '" & cDestino & "') AS MontoCambio, Producto, Contacto, 0 AS Aplicado, 0 AS Incremento, 0 AS Archivado " & _
                                "FROM (SELECT d.Usuario, d.Fecha, d.Codigo + ':' + CAST(d.Secuencia AS varchar(12)) AS CodigoLista, d.Item + ' (Lista: ' + e.Nombre + ')' AS Nombre, e.PrecioOriginal AS LocalPrecio, d.Precio, e.Contacto, d.Secuencia AS Producto " & _
                                "FROM dbo.pre_Listas_Encabezado AS e INNER JOIN dbo.pre_Listas_Detalles AS d ON e.Codigo = d.Codigo AND e.Usuario = d.Usuario WHERE (d.Fecha IS NOT NULL) AND (e.Cuenta = 0)) AS q " & _
                                "WHERE (Fecha BETWEEN '" & Desde & "' AND '" & Hasta & "' );"

                    con.Execute(sqlString)

                    '-----------------------------------------------------------'
                    '                                                           '
                    ' PARTE 5: Verificamos si en el Calendario existen eventos  '
                    '          que afecten el presupuestos y los colocamos...   '
                    '                                                           '
                    '-----------------------------------------------------------'

                    EventosCalendario pre, Desde, Hasta
                end if
            end if

            '-----------------------------------------------------------'
            '                                                           '
            ' Luego de la creación del nuevo presupuesto, verificamos   '
            ' si se especificó un Presupuesto Anterior...               '
            '                                                           '
            '-----------------------------------------------------------'            

            if cAnterior <> "*" then
                '-----------------------------------------------------------'
                '                                                           '
                ' Si hay items sin aplicar en el presupuesto anterior, los  '
                ' movemos al nuevo presupuesto antes de cerrar el anterior  '
                ' de forma automática...                                    '
                '                                                           '
                '-----------------------------------------------------------'

                sqlString = "UPDATE pre_Presupuesto_Detalles " & _
                               "SET Presupuesto = '" & pre & "' " & _
                             "WHERE (Presupuesto = '" & cAnterior & "') " & _
                               "AND (Usuario = '" & usu & "') " & _
                               "AND (Aplicado = 0);"

                con.Execute(sqlString)  

                '-----------------------------------------------------------'
                '                                                           '
                ' Si hay un saldo pendiente o colgante del presupuesto      '
                ' anterior, creamos un registro de cierre para llevar el    '
                ' presupuesto a cero y añadimos ese saldo al nuevo          '
                ' presupuesto...                                            '
                '                                                           '
                '-----------------------------------------------------------'

                restoAnterior = SaldoAnterior(usu, cAnterior)

                if restoAnterior <> 0 then
                    '-----------------------------------------------------------
                    ' Creamos transaccion de cierre en el presupuesto anterior '
                    '-----------------------------------------------------------

                    sqlString = "INSERT INTO pre_Presupuesto_Detalles(Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion," & _
                                                                    " CuentaDestino, MontoDestino, MontoCambio, Aplicado, Incremento, Archivado) " & _
                                    "VALUES ('" & cAnterior & "', '" & usu & "', '" & FechaHoy() & "', " &  HoraHoy() & ", 'PRE-000', " & (-1 * restoAnterior) & ", 'Cierre de Presupuesto - Cartera', " & _
                                    "'SYS-000', " & restoAnterior & ", " & restoAnterior & ", 1, 0, 0);"

                    con.Execute (sqlString)    

                    '------------------------------------------------------
                    ' Cerramos automaticamente el presupuesto anterior... '
                    '------------------------------------------------------

                    sqlString = "UPDATE pre_Presupuesto_Encabezado " & _
                                "SET Estatus = 0 " & _
                                "WHERE Usuario = '" & usu & "' " & _
                                "AND Presupuesto = '" & cAnterior & "';"

                    con.Execute (sqlString)                                                           

                    '------------------------------------------------------------------------------
                    ' Finalmente, creamos el registro de apertura en nuestro nuevo presupuesto... '
                    '------------------------------------------------------------------------------

                    sqlString = "INSERT INTO pre_Presupuesto_Detalles(Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion," & _
                                                                    " CuentaDestino, MontoDestino, MontoCambio, Aplicado, Incremento, Archivado) " & _
                                    "VALUES ('" & pre & "', '" & usu & "', '" & FechaHoy() & "', " &  HoraHoy() & ", 'SYS-000', " & (-1 * restoAnterior) & ", " & _
                                            "'Apertura de Presupuesto - Cartera " & cAnterior & "', " & _
                                    "'PRE-000', " & restoAnterior & ", " & restoAnterior & ", 1, 0, 0);"

                    con.Execute (sqlString)    
                end if


                '--------------------------------------------------'
                ' Ahora verificamos y cerramos los saldos          '
                ' de efectivo e insertamos el registro en el nuevo '
                ' presupuesto                                      '
                '--------------------------------------------------'

                restoAnterior = EfectivoAnterior(usu, cAnterior)

                if restoAnterior <> 0 then
                    '-----------------------------------------------------------
                    ' Creamos transaccion de cierre en el presupuesto anterior '
                    '-----------------------------------------------------------

                    sqlString = "INSERT INTO pre_Presupuesto_Detalles(Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion," & _
                                                                    " CuentaDestino, MontoDestino, MontoCambio, Aplicado, Incremento, Archivado) " & _
                                    "VALUES ('" & cAnterior & "', '" & usu & "', '" & FechaHoy() & "', " &  HoraHoy() & ", 'EF-000', " & (-1 * restoAnterior) & ", 'Cierre de Presupuesto - Efectivo', " & _
                                    "'SYS-000', " & restoAnterior & ", " & restoAnterior & ", 1, 0, 0);"

                    con.Execute (sqlString)    

                    '------------------------------------------------------
                    ' Cerramos automaticamente el presupuesto anterior... '
                    '------------------------------------------------------

                    sqlString = "UPDATE pre_Presupuesto_Encabezado " & _
                                "SET Estatus = 0 " & _
                                "WHERE Usuario = '" & usu & "' " & _
                                "AND Presupuesto = '" & cAnterior & "';"

                    con.Execute (sqlString)                                                           

                    '------------------------------------------------------------------------------
                    ' Finalmente, creamos el registro de apertura en nuestro nuevo presupuesto... '
                    '------------------------------------------------------------------------------

                    sqlString = "INSERT INTO pre_Presupuesto_Detalles(Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion," & _
                                                                    " CuentaDestino, MontoDestino, MontoCambio, Aplicado, Incremento, Archivado) " & _
                                    "VALUES ('" & pre & "', '" & usu & "', '" & FechaHoy() & "', " &  HoraHoy() & ", 'SYS-000', " & (-1 * restoAnterior) & ", " & _
                                            "'Apertura de Presupuesto - Efectivo " & cAnterior & "', " & _
                                    "'EF-000', " & restoAnterior & ", " & restoAnterior & ", 1, 0, 0);"

                    con.Execute (sqlString)    
                end if                
            end if            

            con.close: set con = nothing

           response.redirect "../lista.asp"
        %>    
    </body>
</html>        