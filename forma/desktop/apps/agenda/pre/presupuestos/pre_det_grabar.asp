<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            dim con, sqlString, multi, vinculo, preSiguiente, preSiguienteOperacion
            dim Fecha, Hora, CuentaOrigen, CuentaDestino, preSiguientePresupuesto
            dim Monto, MonedaOrigen, MonedaDestino, MontoCambio
            dim Descripcion, Contacto, Aplicado, Nuevo
            dim usu, pre, llave, mDestino, mCambio
            dim t, v, e, o, sqlStringCopia

            sub Actualizar_Incrementos(Usuario, Presupuesto)
                dim con, t, sqlString

                sqlString = "exec dbo.pre_Presupuestos_Incrementos '" & Usuario & "', '" &  Presupuesto & "'"

                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")
                    con.execute sqlString  
                con.close: set con = nothing            
            end sub

            function Convertir(Monto, MonedaOrigen, MonedaDestino)
                dim con, t, sqlString

                sqlString = "SELECT dbo.Cripto_CambiarMoneda(" & Monto & ", '" & MonedaOrigen & "', '" & MonedaDestino & "') AS Cambio;"

                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")
                set t = con.execute(sqlString)

                Convertir = t("Cambio")

                t.close: set t = nothing
                con.close: set con = nothing
            End Function

            Function FechaSQL(FechaForm)
                dim a, m, d

                FechaSQL = ""

                if not isnull(FechaForm) then
                    d = LEFT(FechaForm, 2)
                    m = MID(FechaForm, 4, 2)
                    a = RIGHT(FechaForm, 4)

                    FechaSQL = a & "-" & m & "-" & d
                end if
            end function

            Function HoraSQL(HoraForm)
                dim h, m

                HoraSQL = ""

                if not isnull(HoraForm) then
                    h = LEFT(HoraForm, 2)
                    m = RIGHT(HoraForm, 2)

                    HoraSQL = h & m
                end if
            end function     

            Function MultiPrecio(Presupuesto)
                dim c, ta

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                set ta = c.Execute("SELECT multiprecio from pre_Presupuesto_Encabezado where (Usuario = '" & Request.Cookies("Usuario")  & "') AND (Presupuesto = '" & presupuesto & "');")

                if not (ta.bof or ta.eof) then
                    MultiPrecio = ta("multiprecio")
                else
                    MultiPrecio = 0
                end if

                ta.close: set ta = nothing
                c.close: set c = nothing
            end function

            Function ListaMonto(Lista, NuevoLocalMonetario)
                dim c, ta, sqlString, CodigoLista

                CodigoLista = Right(Lista, (Len(Lista) -3))

                sqlString = "SELECT dbo.pre_TotalizarLista('" & Request.Cookies("Usuario") & "', '" & CodigoLista & "', '" & NuevoLocalMonetario & "') AS Total;"
                ListaMonto = 0

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                set ta = c.Execute(sqlString)

                if not (ta.bof or ta.eof) then
                    ListaMonto = ta("Total")
                end if

                ta.close: set ta = nothing
                c.close: set c = nothing   
            end Function

            Function ListaContacto(Lista)
                dim c, ta, sqlString, CodigoLista

                ListaContacto = NULL

                if InStr(Lista,"*L:") then  
                    ListaContacto = Trim(Lista)
                    CodigoLista = RIGHT(ListaContacto, (len(ListaContacto) - 3))

                    sqlString = "select Contacto " & _
                                "from pre_Listas_Encabezado " & _
                                "where codigo = '" & CodigoLista & "' " & _
                                "and usuario ='" & Request.Cookies("Usuario") & "' " & _
                                "and cuenta = 1;"

                    set c = server.CreateObject("ADODB.Connection")
                    c.open Application("Conn")
                    set ta = c.Execute(sqlString)

                    if not (ta.bof or ta.eof) then
                        ListaContacto = ta("Contacto")
                    else
                        ListaContacto = NULL
                    end if

                    ta.close: set ta = nothing
                    c.close: set c = nothing 
                end if
            end Function    

            function LimpiarApostrofes(valor)
                LimpiarApostrofes = Replace(valor,"'","´")
                LimpiarApostrofes = Replace(LimpiarApostrofes, "<","&#11013;")
                LimpiarApostrofes = Replace(LimpiarApostrofes, ">","&#11157;")
            end function       

            Sub CopiarRegistro(Llave, NuevoPresupuesto)
                dim c, ta, sqlInsert

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    sqlInsert = "INSERT INTO pre_Presupuesto_Detalles(Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, Aplicado, Contacto, ItemLista, NotaPre, NotaDonde, Nota) " & _
                                "SELECT '" & NuevoPresupuesto & "', Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, Aplicado, Contacto, ItemLista, NotaPre, NotaDonde, Nota " & _
                                "FROM pre_Presupuesto_Detalles " & _ 
                                "WHERE (Llave = " & Llave  & ");"
                    c.execute(sqlInsert)
                c.close: set c = nothing            
            End Sub
        %>
    </head>

    <body>
        <%
            '
            ' Esta pagina graba la transaccion en la base de datos
            '

            usu = Request.form("Usuario")
            pre = Request.form("Presupuesto")
            d = Request.Form("d")
            t = Request.Form("t")
            v = Request.Form("v")
            e = Request.Form("e")
            o = Request.Form("o")

            vinculo = "pre_det_editar.asp?p=" & pre & "&d=" & d & "&v=" & v & "&t=" & t & "&e=" & e & "&o=" & o                        

            '
            ' Leemos el formulario
            '
            Fecha = FechaSQL(Request.form("txt_fecha"))
            Hora = HoraSQL( Request.form("txt_hora"))
            CuentaOrigen = Request.form("CuentaOrigen")
            CuentaDestino = Request.form("CuentaDestino")
            Monto = Replace(Request.form("Monto"), "," ,"")
            MontoCambio = Replace(Request.Form("txtMontoCambio"), ",", "")
            MonedaOrigen = Request.form("MonedaOrigen")
            MonedaDestino = Request.form("MonedaDestino")
            Descripcion = LimpiarApostrofes(Request.form("Descripcion"))
            Contacto = Request.form("Contacto")
            Aplicado = Request.form("Aplicado")   
            preSiguiente = Request.form("preSiguiente")     

            if preSiguiente = "*" then
                preSiguienteOperacion = "*"
            else
                preSiguienteOperacion = LEFT(preSiguiente, 1)
                preSiguientePresupuesto = RIGHT(preSiguiente, (len(preSiguiente) - 2))
            end if

            llave = Request.form("Llave")
            nuevo = Request.form("Nuevo")
            multi = MultiPrecio(pre)

            if multi = "" then
                if MonedaOrigen <> MonedaDestino then
                    multi = 1
                else
                    multi = 0
                end if
            end if

            '
            ' Verificar Listas...
            '

            if InStr(CuentaDestino, "*L:") then
                '
                ' Es una Lista convertida en Cuenta...
                '
                Monto = ListaMonto(CuentaDestino, MonedaOrigen)
                MontoCambio = ListaMonto(CuentaDestino, MonedaDestino)

                Contacto = ListaContacto(CuentaOrigen)
                mDestino = monto
                monto = (-1) * Monto
            else            
                mDestino = monto
                monto = (-1) * Monto
            end if

            if Contacto = "" then Contacto = NULL

            '
            ' Creamos la cadena de SQL
            '

            if Nuevo = 1 then
                '
                ' CREAMOS EL REGISTRO
                '
                sqlString = "INSERT INTO pre_Presupuesto_Detalles(Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, Aplicado, Contacto) " & _
                            "VALUES ('" & pre & "', '" & usu & "', '" & Fecha & "', " & Hora & ", '" & CuentaOrigen & "'," & monto & ", '" & Descripcion & "'" & _
                                    ", '" & CuentaDestino & "', " & mDestino & ", " & MontoCambio & ", " & Aplicado & ", '" & Contacto & "');"
            else
                '
                ' UPDATE
                '
                sqlString = "UPDATE pre_Presupuesto_Detalles " & _
                            "SET Fecha = '" & Fecha & "', " & _
                               " Hora = " & Hora & ", " & _
                               " CuentaOrigen = '" & CuentaOrigen & "', " & _
                               " MontoOrigen = " & monto & ", " & _
                               " Descripcion = '" & Descripcion & "', " & _
                               " CuentaDestino = '" & CuentaDestino & "', " & _
                               " MontoDestino = " & mDestino & ", " & _
                               " MontoCambio = " & MontoCambio & ", " & _
                               " Aplicado = " & Aplicado & ", " & _
                               " Contacto = '" & Contacto & "'" 

                if preSiguienteOperacion = "M" then
                    sqlString = sqlString & ", Presupuesto = '" & preSiguientePresupuesto & "' " 
                end if

                sqlString = sqlString & " WHERE (Llave = " & Llave & ");"
            end if

            '
            ' Ejecutamos el comando y volvemos al presupuesto
            '
            response.write sqlString

            set con = server.CreateObject("ADODB.Connection")
            con.open Application("Conn")
                con.execute(sqlString)

                if preSiguienteOperacion = "C" then
                    '
                    ' Copiamos la Transaccion si se ha solicitado
                    '

                    CopiarRegistro Llave, preSiguientePresupuesto
                end if                
            con.close: set con = nothing

            '
            ' Actualizamos los Incrementos si el Presupuesto es un Modelo...
            '
            Actualizar_Incrementos usu, pre

            '
            ' Volvemos al Presupuesto
            '
            response.redirect vinculo
        %>    
    </body>
</html>        