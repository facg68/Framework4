<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Cerrar Presupuesto" %>
        <title><%= PageTitle %></title>

        <%
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

            function TransaccionesSinAplicar(Usuario, Presupuesto)            
                dim cc, tt, sqlString

                TransaccionesSinAplicar = 0

                sqlString = "SELECT COUNT(*) AS Cuantos " & _
                            "FROM dbo.pre_Presupuesto_Detalles " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Presupuesto = '" & Presupuesto & "') " & _
                            "AND (Aplicado = 0);"
                
                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                    set tt = cc.execute(sqlString)
                        if not (tt.bof or tt.eof) then
                            TransaccionesSinAplicar = tt("Cuantos")
                        end if
                    tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function
        %>        
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <div class="page-title-bar">
            <%= PageTitle %>
        </div>

        <main>
            <%
                dim Usuario, Presupuesto, Saldos, Transacciones

                Usuario = request.Cookies("Usuario")
                Presupuesto = request.QueryString("pre") 

                Saldos = SaldoActual(Usuario, Presupuesto) + SaldoActualE(Usuario, Presupuesto)
                Transacciones = TransaccionesSinAplicar(Usuario, Presupuesto)

                if (Saldos > 0) then    
                    '
                    ' Seleccionamos el Presupuesto al que asignaremos el saldo de cierre...
                    '
                    response.redirect "pre_det_cerrar2.asp?p=" & Presupuesto    
                else
                    response.redirect "pre_det_cerrar4.asp?p=" & Presupuesto & "&np=*"
                end if
            %>
        </main>

        <footer class="footer-contextual">
            <button class="footer-button" type="button" aria-label="Volver" onclick="volver()">
                <i class="fas fa-undo-alt"></i>
            </button>
        </footer>

        <script>
            function volver() {
                history.back();
            }
        </script>

        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>