<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Cerrar Presupuesto" %>
        <title><%= PageTitle %></title>

        <%
            function FechaServer()
                dim f, dia, mes, amo

                f = DateAdd("h", 2, Now()) 

                dia = Day(f)
                mes = Month(f)
                amo = Year(f)

                FechaServer = amo & "-" &  Right("00" & Mes, 2) & "-" & right("00" & dia, 2)
            end function

            function HoraServer()
                dim h, m, cadena, f

                f = DateAdd("h", 2, Now()) 

                h = hour(f)
                m = minute(f)

                cadena = h & right("00" & m, 2)
                HoraServer = cint(cadena)
            end function

            function SaldoActual(Usuario, Presupuesto)
                dim c, ta, sqlString, db, cr

                sqlString = "SELECT SUM(iif(CuentaOrigen = 'PRE-000' AND Aplicado = 1, MontoOrigen,0)) AS DB, " & _
                                  " SUM(iif(CuentaDestino = 'PRE-000' AND Aplicado = 1,MontoDestino,0)) AS CR " & _
                              "FROM pre_Presupuesto_Detalles " & _
                             "WHERE (Usuario = '" & Usuario & "') " & _
                               "AND (Presupuesto = '" & Presupuesto & "')"

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set ta = c.Execute(sqlString)
                        if not (ta.bof or ta.eof) then
                            cr = ta("CR")
                            db = ta("DB")

                            if isnull(cr) or cr = "" then cr = 0
                            if isnull(db) or db = "" then db = 0

                            SaldoActual = db + cr        
                        else
                            SaldoActual = 0.00
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
            end Function 

            function SaldoActualE(Usuario, Presupuesto)
                dim c, ta, sqlString, db, cr

                sqlString = "SELECT SUM(iif(CuentaOrigen = 'EF-000' AND Aplicado = 1, MontoOrigen,0)) AS DB, " & _
                                        " SUM(iif(CuentaDestino = 'EF-000' AND Aplicado = 1,MontoDestino,0)) AS CR " & _
                                        "FROM pre_Presupuesto_Detalles " & _
                                        "WHERE (Usuario = '" & Usuario & "') " & _
                                        "AND (Presupuesto = '" & Presupuesto & "')"

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set ta = c.Execute(sqlString)
                        if not (ta.bof or ta.eof) then
                            cr = ta("CR")
                            db = ta("DB")

                            if isnull(cr) or cr = "" then cr = 0
                            if isnull(db) or db = "" then db = 0

                            SaldoActualE = db + cr        
                        else
                            SaldoActualE = 0.00
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
            end Function   

            function Cambio(Usuario, Presupuesto, MontoDestino)
                dim cc, tt, sqlString, tc

                sqlString = "SELECT MonedaOrigen, MonedaDestino " & _
                            "FROM dbo.pre_Presupuesto_Encabezado " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Presupuesto = '" & Presupuesto & "');" 
                
                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                    set tt = cc.execute(sqlString)
                        if tt("MonedaOrigen") <> tt("MonedaDestino") then
                            sqlString = "SELECT dbo.Cripto_CambiarMoneda(" & MontoDestino & ", '" & tt("MonedaOrigen") & "', '" & tt("MonedaDestino") & "') AS Cambio;"

                            set tc = cc.execute(sqlString)
                                Cambio = tc("Cambio")
                            tc.close: set tc = nothing
                        else
                            Cambio = MontoDestino
                        end if
                    tt.close: set tt = nothing
                cc.close: set cc = nothing      
            end function

            sub GenerarCierre(Usuario, Presupuesto, nPresupuesto)
                dim cc, sqlString, MontoOrigen, MontoDestino, MontoCambio

                MontoDestino = SaldoActual(Usuario, Presupuesto)

                if MontoDestino > 0.00 then     
                    MontoOrigen = (-1 * MontoDestino)
                    MontoCambio = Cambio(Usuario, Presupuesto, MontoDestino)

                    sqlString = "INSERT INTO pre_Presupuesto_Detalles(Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, Aplicado) " & _
                                "VALUES ('" & Presupuesto & "', '" & Usuario & "', '" & FechaServer() & "', " & HoraServer() & ", 'PRE-000'," & MontoOrigen & ", 'Cierre de Presupuesto - Cartera', " & _
                                        "'SYS-000', " & MontoDestino & ", " & MontoCambio & ", 1);"
                    
                    set cc = Server.CreateObject("ADODB.Connection")

                    cc.open Application("Conn")
                        cc.execute(sqlString)
                    cc.close: set cc = nothing

                    if nPresupuesto <> "*" then
                        GenerarApertura Usuario, nPresupuesto, MontoDestino
                    end if
                end if
            end sub

            sub GenerarApertura(Usuario, nPresupuesto, Monto)
                dim cc, sqlString, MontoOrigen, MontoCambio

                if Monto > 0.00 then     
                    MontoOrigen = (-1 * Monto)
                    MontoCambio = Cambio(Usuario, nPresupuesto, Monto)

                    sqlString = "INSERT INTO pre_Presupuesto_Detalles(Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, Aplicado) " & _
                                "VALUES ('" & nPresupuesto & "', '" & Usuario & "', '" & FechaServer() & "', " & HoraServer() & ", 'SYS-000'," & MontoOrigen & ", 'Apertura de Presupuesto - Cartera', " & _
                                        "'PRE-000', " & Monto & ", " & MontoCambio & ", 1);"
                    
                    set cc = Server.CreateObject("ADODB.Connection")

                    cc.open Application("Conn")
                        cc.execute(sqlString)
                    cc.close: set cc = nothing
                end if
            end sub

            sub GenerarCierreEfectivo(Usuario, Presupuesto, nPresupuesto)
                dim cc, sqlString, MontoOrigen, MontoDestino, MontoCambio

                MontoDestino = SaldoActualE(Usuario, Presupuesto)

                if MontoDestino > 0.00 then        
                    MontoOrigen = (-1 * MontoDestino)
                    MontoCambio = Cambio(Usuario, Presupuesto, MontoDestino)

                    sqlString = "INSERT INTO pre_Presupuesto_Detalles(Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, Aplicado) " & _
                                "VALUES ('" & Presupuesto & "', '" & Usuario & "', '" & FechaServer() & "', " & HoraServer() & ", 'EF-000'," & MontoOrigen & ", 'Cierre de Presupuesto - Efectivo', " & _
                                        "'SYS-000', " & MontoDestino & ", " & MontoCambio & ", 1);"
                    
                    set cc = Server.CreateObject("ADODB.Connection")

                    cc.open Application("Conn")
                        cc.execute(sqlString)
                    cc.close: set cc = nothing

                    if nPresupuesto <> "*" then
                        GenerarAperturaEfectivo Usuario, nPresupuesto, MontoDestino
                    end if
                end if
            end sub

            sub GenerarAperturaEfectivo(Usuario, nPresupuesto, Monto)
                dim cc, sqlString, MontoOrigen, MontoCambio

                if Monto > 0.00 then        
                    MontoOrigen = (-1 * Monto)
                    MontoCambio = Cambio(Usuario, Presupuesto, Monto)

                    sqlString = "INSERT INTO pre_Presupuesto_Detalles(Presupuesto, Usuario, Fecha, Hora, CuentaOrigen, MontoOrigen, Descripcion, CuentaDestino, MontoDestino, MontoCambio, Aplicado) " & _
                                "VALUES ('" & nPresupuesto & "', '" & Usuario & "', '" & FechaServer() & "', " & HoraServer() & ", 'SYS-000'," & MontoOrigen & ", 'Apertura de Presupuesto - Efectivo', " & _
                                        "'EF-000', " & Monto & ", " & MontoCambio & ", 1);"
                    
                    set cc = Server.CreateObject("ADODB.Connection")

                    cc.open Application("Conn")
                        cc.execute(sqlString)
                    cc.close: set cc = nothing
                end if
            end sub      

            function TieneTransacciones(Usuario, Presupuesto)
                dim cc, tt, sqlString

                sqlString = "SELECT ISNULL(COUNT(*), 0) AS Cuantos " & _
                            "FROM pre_Presupuesto_Detalles " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Presupuesto = '" & Presupuesto & "') " & _
                            "AND (Aplicado = 0);"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                    set tt = cc.Execute(sqlString)
                        TieneTransacciones = tt("Cuantos")				
                    tt.close: set tt = nothing
                cc.close: set cc = nothing			
            end function

            Sub MoverTransacciones(Usuario, Presupuesto, nPresupuesto)
                dim cc, sqlString, MontoOrigen, MontoDestino, MontoCambio

                if nPresupuesto <> "*" then
                    sqlString = "UPDATE pre_Presupuesto_Detalles " & _
                                "SET Presupuesto = '" & nPresupuesto & "' " & _
                                "WHERE (Usuario = '" & Usuario & "') " & _
                                "AND (Presupuesto = '" & Presupuesto & "') " & _
                                "AND (Aplicado = 0);"

                    set cc = Server.CreateObject("ADODB.Connection")
                    cc.open Application("Conn")
                        cc.execute(sqlString)
                    cc.close: set cc = nothing
                end if
            end sub
        %>        
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <div class="page-title-bar">
            <%= PageTitle %>
        </div>

        <main>
            <div class="contenedor">
                <%
                    dim c, sqlString, Usuario, Presupuesto, nPresupuesto

                    Usuario = request.Cookies("Usuario")
                    Presupuesto = request.QueryString("p") 
                    nPresupuesto = request.QueryString("np")

                    if (cDbl(SaldoActual(Usuario, Presupuesto)) + cDbl(SaldoActualE(Usuario, Presupuesto)) ) > cDbl(0.00) then
                        GenerarCierre Usuario, Presupuesto, nPresupuesto
                        GenerarCierreEfectivo Usuario, Presupuesto, nPresupuesto
                    end if

                    MoverTransacciones Usuario, Presupuesto, nPresupuesto

                    sqlString = "UPDATE pre_Presupuesto_Encabezado " & _
                                   "SET Estatus = 0 " & _
                                 "WHERE (Usuario = '" & Usuario & "') " & _
                                   "AND (Presupuesto = '" & Presupuesto & "');"

                    set c = server.CreateObject("ADODB.Connection")
                    c.open Application("Conn")
                        c.execute (sqlString)
                    c.close: set c = nothing

                    response.redirect "../lista.asp"   
                %>  
            <div>
        </main>

        <footer class="footer-contextual">
            <button class="footer-button" type="button" aria-label="Home" onclick="irA('/forma/mobile')">
                <i class="fa-solid fa-house"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Volver" onclick="volver()">
                <i class="fas fa-undo-alt"></i>
            </button>
        </footer>

        <script>
            function volver() {
                history.back();
            }

            function irA(vinculo) {
                window.location.href = vinculo;
            }            
        </script>

        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>