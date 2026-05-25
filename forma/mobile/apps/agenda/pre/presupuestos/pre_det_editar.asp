<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Editar Presupuesto" %>
        <title><%= PageTitle %></title>

        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0090"
            SysLockOut


            '
            ' Global  Init()
            '

            dim con, t, l, sqlCommand, sqlString, usu, pre, est, cuantos, sw
            dim suma, mPrecio, tipo, subTotal1, subTotal2, ver, dia, colCampo
            dim aClass, wFecha, wDescripcion, wMontoOrigen, ClaseTabla, clasePre
            dim wMontoCambio, wContacto, wAplicado, wAcciones
            dim redirectTipo, redirectEstatus, redirectOrdenamiento
            dim labelLocalOriginal, labelLocalCambio

            usu = Request.Cookies("usuario")
            pre = Request.QueryString("pre")

            if pre = "" then pre = request.Cookies("pre_det_pre")
            if pre = "" then response.redirect "../lista.asp"

            response.Cookies("pre_det_pre") = pre

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")    

            ' Funciones y Procedimientos ------------------------------------------------------------------------------------
                function NombrePresupuesto()
                    dim ta

                    set ta = con.Execute("SELECT nombre from pre_Presupuesto_Encabezado " & _
                                        "WHERE (Usuario = '" & usu & "') " & _
                                            "AND (Presupuesto = '" & pre & "');")

                        if not (ta.bof or ta.eof) then
                            NombrePresupuesto = ta("nombre")
                        else
                            NombrePresupuesto = ""
                        end if
                    ta.close: set ta = nothing
                end Function

                function TipoPresupuesto()
                    dim ta

                    set ta = con.Execute("SELECT tipo from pre_Presupuesto_Encabezado " & _
                                        "WHERE (Usuario = '" & usu & "') " & _
                                            "AND (Presupuesto = '" & pre & "');")

                        if not (ta.bof or ta.eof) then
                            TipoPresupuesto = ta("tipo")
                        else
                            TipoPresupuesto = ""
                        end if
                    ta.close: set ta = nothing
                end Function      

                function EstatusPresupuesto()
                    dim ta

                    set ta = con.Execute("SELECT estatus from pre_Presupuesto_Encabezado " & _
                                        "WHERE (Usuario = '" & usu & "') " & _
                                            "AND (Presupuesto = '" & pre & "');")

                        if not (ta.bof or ta.eof) then
                            EstatusPresupuesto = ta("estatus")
                        else
                            EstatusPresupuesto = 0
                        end if
                    ta.close: set ta = nothing
                end Function  

                Function MultiPrecio()
                    dim ta

                    set ta = con.Execute("SELECT multiprecio from pre_Presupuesto_Encabezado " & _
                                        "WHERE (Usuario = '" & usu & "') " & _
                                            "AND (Presupuesto = '" & pre & "');")

                        if not (ta.bof or ta.eof) then
                            MultiPrecio = ta("multiprecio")
                        else
                            MultiPrecio = 0
                        end if
                    ta.close: set ta = nothing
                end function   

                function LocalMontoOriginal()
                    dim ta

                    set ta = con.Execute("SELECT MonedaOrigen from pre_Presupuesto_Encabezado " & _
                                        "WHERE (Usuario = '" & usu & "') " & _
                                            "AND (Presupuesto = '" & pre & "');")

                        if not (ta.bof or ta.eof) then
                            LocalMontoOriginal = ta("MonedaOrigen")
                        else
                            LocalMontoOriginal = ""
                        end if
                    ta.close: set ta = nothing
                end function

                function LocalMontoCambio()
                    dim c, ta

                    set ta = con.Execute("SELECT MonedaDestino FROM pre_Presupuesto_Encabezado " & _
                                        "WHERE (Usuario = '" & usu & "') " & _
                                            "AND (Presupuesto = '" & pre & "');")

                        if not (ta.bof or ta.eof) then
                            LocalMontoCambio = ta("MonedaDestino")
                        else
                            LocalMontoCambio = ""
                        end if
                    ta.close: set ta = nothing
                end function  

                function SaldoPresupuestoCartera()
                    dim ta, sqlString, db, cr

                    sqlString = "SELECT SUM(iif(CuentaOrigen = 'PRE-000' AND Aplicado = 1, MontoOrigen,0)) AS DB, " & _
                                            " SUM(iif(CuentaDestino = 'PRE-000' AND Aplicado = 1,MontoDestino,0)) AS CR " & _
                                            "FROM pre_Presupuesto_Detalles " & _
                                            "WHERE (Usuario = '" & usu & "') " & _
                                            "AND (Presupuesto = '" & pre & "')"

                    set ta = con.Execute(sqlString)
                        if not (ta.bof or ta.eof) then
                            cr = ta("CR")
                            db = ta("DB")

                            if isnull(cr) or cr = "" then cr = 0
                            if isnull(db) or db = "" then db = 0

                            SaldoPresupuestoCartera = db + cr        
                        else
                            SaldoPresupuestoCartera = 0.00
                        end if
                    ta.close: set ta = nothing
                end Function   

                function SaldoPresupuestoEfectivo()
                    dim ta, sqlString, db, cr

                    sqlString = "SELECT SUM(iif(CuentaOrigen = 'EF-000' AND Aplicado = 1, MontoOrigen,0)) AS DB, " & _
                                            " SUM(iif(CuentaDestino = 'EF-000' AND Aplicado = 1,MontoDestino,0)) AS CR " & _
                                            "FROM pre_Presupuesto_Detalles " & _
                                            "WHERE (Usuario = '" & usu & "') " & _
                                            "AND (Presupuesto = '" & pre & "')"

                    set ta = con.Execute(sqlString)
                        if not (ta.bof or ta.eof) then
                            cr = ta("CR")
                            db = ta("DB")

                            if isnull(cr) or cr = "" then cr = 0
                            if isnull(db) or db = "" then db = 0

                            SaldoPresupuestoEfectivo = db + cr        
                        else
                            SaldoPresupuestoEfectivo = 0.00
                        end if
                    ta.close: set ta = nothing
                end Function           

                function SaldoPresupuestoCarteraTotal()
                    dim c, ta, sqlString, db, cr

                    sqlString = "SELECT SUM(iif(CuentaOrigen = 'PRE-000', MontoOrigen,0)) AS DB, " & _
                                            " SUM(iif(CuentaDestino = 'PRE-000',MontoDestino,0)) AS CR " & _
                                            "FROM pre_Presupuesto_Detalles " & _
                                            "WHERE (Usuario = '" & usu & "') " & _
                                            "AND (Presupuesto = '" & pre & "')"

                    set ta = con.Execute(sqlString)
                        if not (ta.bof or ta.eof) then
                            cr = ta("CR")
                            db = ta("DB")

                            if isnull(cr) or cr = "" then cr = 0
                            if isnull(db) or db = "" then db = 0

                            SaldoPresupuestoCarteraTotal = db + cr
                        else
                            SaldoPresupuestoCarteraTotal = 0.00
                        end if
                    ta.close: set ta = nothing
                end Function     

                function SaldoPresupuestoEfectivoTotal()
                    dim c, ta, sqlString, db, cr

                    sqlString = "SELECT SUM(iif(CuentaOrigen = 'EF-000', MontoOrigen,0)) AS DB, " & _
                                            " SUM(iif(CuentaDestino = 'EF-000',MontoDestino,0)) AS CR " & _
                                            "FROM pre_Presupuesto_Detalles " & _
                                            "WHERE (Usuario = '" & usu & "') " & _
                                            "AND (Presupuesto = '" & pre & "')"

                    set ta = con.Execute(sqlString)
                        if not (ta.bof or ta.eof) then
                            cr = ta("CR")
                            db = ta("DB")

                            if isnull(cr) or cr = "" then cr = 0
                            if isnull(db) or db = "" then db = 0

                            SaldoPresupuestoEfectivoTotal = db + cr
                        else
                            SaldoPresupuestoEfectivoTotal = 0.00
                        end if
                    ta.close: set ta = nothing
                end Function                    
                
                function HtmlResumenCuentas()
                    dim html

                    html = "<div class='cuentas-resumen'>"
                        html = html & CuentaResumen("Cartera", SaldoPresupuestoCartera())
                        html = html & CuentaResumen("Cartera Presupuestada", SaldoPresupuestoCarteraTotal())
                        html = html & CuentaResumen("Efectivo", SaldoPresupuestoEfectivo())
                        html = html & CuentaResumen("Efectivo Presupuestado", SaldoPresupuestoEfectivoTotal())
                    html = html & "</div>"

                    HtmlResumenCuentas = html
                end function

                function CuentaResumen(nombre, saldo)
                    CuentaResumen = _
                        "<div class='cuenta-row'>" & _
                            "<div class='cuenta-nombre'>" & nombre & "</div>" & _
                            "<div class='cuenta-saldo-linea'>" & _
                                "<span class='saldo-label'>Saldo:</span>" & _
                                "<span class='cuenta-saldo'>" & FormatNumber(saldo, 2) & "</span>" & _
                            "</div>" & _
                        "</div>"
                end function

                function claseVinculo(CuentaOrigen, CuentaDestino, Monto, Aplicada) 
                    claseVinculo = "otros_"

                    if cuentaOrigen = "PRE-000" then claseVinculo = "cartera_DB_"
                    if cuentaDestino = "PRE-000" then claseVinculo = "cartera_CR_"
                    
                    if cuentaOrigen = "EF-000" then claseVinculo = "efectivo_DB_"
                    if cuentaDestino = "EF-000" then claseVinculo = "efectivo_CR_" 

                    if (cuentaOrigen = CuentaDestino) or (Monto = 0) then claseVinculo = "otros_"

                    claseVinculo = claseVinculo & Aplicada
                end function                                      
            ' ---------------------------------------------------------------------------------------------------------------

        %>   

        <style>
            /* Clases Base --------------------------- */
                .presupuesto-item {
                    display: flex;
                    align-items: center;
                    gap: 14px;
                    width: 100%;
                    padding: 14px 12px;
                    text-decoration: none;
                    color: inherit;
                    border-bottom: 1px solid rgba(0,0,0,0.08);
                }    
                
                .pre-verde { background-color: #154927; }
                .pre-azul { background-color: #475eb0; }
                .pre-morado { background-color: #6b3194; }             
                .presupuesto-item { padding-left: 16px; }            
                .presupuesto-item:active { background: rgba(0,0,0,0.06); }

                .presupuesto-info {
                    flex: 1;
                    min-width: 0;
                }

                .presupuesto-nombre {
                    font-family: 'Ruda Bold';
                    font-size: 1.15rem;
                    white-space: normal;        
                    overflow: visible;          
                    text-overflow: clip;        
                    word-break: break-word;     
                }            

                .presupuesto-datos {
                    font-family: 'Ruda Bold';
                    font-size: 1rem;
                }  
            /*---------------------------------------- */

            /* Cuadro de Resumen de Cuentas ---------- */
                .cuentas-resumen {
                    width: 100%;
                }

                .cuenta-row {
                    padding: 8px 0;
                    border-bottom: 1px solid #e2e2e2;
                }

                .cuenta-nombre {
                    font-weight: 600;
                    font-size: 0.95rem;
                    margin-bottom: 2px;
                }

                .cuenta-saldo-linea {
                    font-size: 0.9rem;
                }

                .saldo-label {
                    color: #666;
                }

                .cuenta-saldo {
                    font-weight: 700;
                    color: #1e4f63;
                }
            /* --------------------------------------- */

            /* SWAL de Resumen de Cuentas ------------ */
                .swal2-popup {
                    width: 95% !important;
                    max-width: 420px !important;   /* más compacto en móvil */
                    max-height: 65vh !important;   /* antes 85vh */
                    display: flex !important;
                    flex-direction: column;
                    padding: 14px !important;
                }

                .swal2-html-container {
                    flex: 1 1 auto;
                    overflow-y: auto !important;
                    margin: 0 !important;
                }

                .swal2-actions {
                    flex-shrink: 0;
                    margin-top: 8px !important;
                }
            /* --------------------------------------- */

            /* Tipos de Transacciones */
                
                .cartera_CR_0 { color: rgb(3, 80, 11); }
                .cartera_CR_1 { color: rgb(123, 150, 126); }   

                .cartera_DB_0 { color: rgb(96, 6, 6); }
                .cartera_DB_1 { color: rgb(155, 106, 106); }
                
                .efectivo_CR_0 { color: rgb(2, 35, 113); }
                .efectivo_CR_1 { color: rgb(97, 112, 146) }
                
                .efectivo_DB_0 { color: rgb(72, 9, 91); }
                .efectivo_DB_1 { color: rgb(134, 101, 145); }
                
                .otros_0 { color:rgb(0, 0, 0); }
                .otros_1 { color:rgb(124, 123, 123); } 

                .fondo_0 { background-color: rgb(255, 255, 255); }
                .fondo_1 { background-color: rgb(243, 243, 243); }

            /* Fin - Tipos de Transacciones */               
        </style>             
    </head>

    <body plus= ".presupuesto-item : plus_telPreDetalles; ">
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->   

        <% 
            ' Preparar Datos ----------------------------------------------------------------------------------------------------

                est = EstatusPresupuesto()
                ver = request.Cookies("pre_det_ver")
                dia = request.Cookies("pre_det_dia")          

                mPrecio = MultiPrecio()
                tipo = TipoPresupuesto()

                if tipo = "P" then
                    select case est
                        case "0"
                            clasePre = "azul"
                        case "1"
                            clasePre = "verde"
                        case "2"
                            clasePre = "verde"
                    end select
                else
                    clasePre = "morado"
                end if

                if dia = "" then dia = "*"
                if ver = "" then ver = "0"
                if est = 0  then ver = "*"

                sqlCommand = "SELECT Llave, Usuario, Presupuesto, TipoPre, Estatus, FechaTransaccion, Descripcion, " & _
                                   " MontoDestino, MontoCambio, Aplicado, Archivado, MultiPrecio, SimboloOrigen, SimboloDestino, " & _
                                   " CuentaOrigen, CuentaDestino, NombreCuentaOrigen, NombreCuentaDestino " & _
                               "FROM pre_TransaccionesResumidas " & _
                              "WHERE (Usuario = '" & usu & "') AND (Presupuesto = '" & pre & "') " 

                if ver = "1" then sqlCommand = sqlCommand & "AND (Aplicado = 1) "
                if ver = "0" then sqlCommand = sqlCommand & "AND (Aplicado = 0) "
                if dia <> "*" then sqlCommand = sqlCommand & "AND (Fecha = '" & dia & "') "

                sqlCommand = sqlCommand & "ORDER BY Fecha, Hora;"   

            '--------------------------------------------------------------------------------------------------------------------   

            set t = con.Execute(sqlCommand)
        %>  

        <div class="page-title-bar pre-<%= clasePre %>">
            <div class="title"><%= NombrePresupuesto() %></div>
        </div>

        <main>
            <%
                if not (t.bof or t.eof) then
                    cuantos = 0

                    Do
                        cuantos = cuantos + 1

                        if (est = 0) then
                            '
                            ' El presupuesto está cerrado, pero se ha abierto para consulta...
                            ' Presentamos los detalles con colores vivos (tipo "abierto")
                            ' para ver mejor los datos. Igual no puede ser modificado.
                            '
                            aClass = claseVinculo(t("CuentaOrigen"), t("CuentaDestino"), t("MontoDestino"), 0)
                            fondo_a = "fondo_0"
                        else
                            aClass = claseVinculo(t("CuentaOrigen"), t("CuentaDestino"), t("MontoDestino"), t("Aplicado"))
                            fondo_a = "fondo_" & t("Aplicado")
                        end if                   

                        %>
                            <a class="presupuesto-item <%= fondo_a %>" id="<%= t("Llave") %>"
                               data-registro="<%= t("Llave") %>"
                               data-presupuesto="<%= t("Presupuesto") %>"
                               data-estatus_pre="<%= t("Estatus") %>"
                               data-tipo="<%= t("TipoPre") %>"
                               data-estatus="<%= t("Aplicado") %>"
                               data-archivado="<%= t("Archivado") %>"                               
                            >
                                <div class="presupuesto-info">
                                    <div class="presupuesto-nombre <%= aClass %>">
                                        <%= t("Descripcion") %>
                                    </div>

                                    <div class="presupuesto-datos <%= aClass %>">
                                        <%
                                            response.write t("FechaTransaccion") & "<br/>"
                                            response.write "De: " & t("NombreCuentaOrigen") & "<br/>"
                                            response.write "A: " & t("NombreCuentaDestino") & "<br/>"

                                            if mPrecio = 1 then
                                                response.write "Monto Origen: (" & t("SimboloOrigen") & ") " & FormatNumber(t("MontoDestino"), 2, -1, 0, -1) & "<br/>"
                                                response.write "Cambio: (" & t("SimboloDestino") & ") " & FormatNumber(t("MontoCambio"), 2, -1, 0, -1)
                                            else
                                                response.write "Monto: (" & t("SimboloDestino") & ") " & FormatNumber(t("MontoCambio"), 2, -1, 0, -1)
                                            end if
                                        %>
                                    </div>
                                </div>
                            </a>
                        <%

                        t.movenext
                    loop until t.eof
                end if
            %>
        </main>

        <footer class="footer-contextual">
            <button class="footer-button" type="button" aria-label="Home" onclick="irA('/forma/mobile')">
                <i class="fa-solid fa-house"></i>
            </button>
                    
            <button class="footer-button" type="button" aria-label="Volver" onclick="volver()">
                <i class="fas fa-undo-alt"></i>
            </button>

            <button class="footer-button" aria-label="Nuevo" onclick="nuevo()">
                <i class="fa-solid fa-plus"></i>
            </button>            

            <button class="footer-button" aria-label="Cuentas" onclick="cuentas()">
                <i class="fa-solid fa-wallet"></i>
            </button>            

            <button class="footer-button" aria-label="Filtrar" onclick="filtro()">
                <i class="fa-solid fa-list-check"></i>
            </button>                
        </footer>

        <script>
            function volver() {
                history.back();
            }

            function irA(vinculo) {
                window.location.href = vinculo;
            }            

            function nuevo() {
                var vinculo = "pre_det_tra_editar.asp?pre=<%= pre %>&registro=0";
                window.location.href = vinculo;
            }

            function filtro() {
                var vinculo = "filtrar.asp";
                window.location.href = vinculo;
            }         
            
            function cuentas() {
                Swal.fire({
                    html: `<%= Replace(HtmlResumenCuentas(), "`", "\`") %>`,
                    width: 600,
                    showConfirmButton: true,
                    confirmButtonText: 'Cerrar',
                    confirmButtonColor: '#2f6fb2',
                    customClass: {
                        popup: 'swal2-border-radius'
                    }
                });
            }               
        </script>

        <% con.close: set con = nothing %>

        <!-- #include virtual = "/forma/plus/telPreDetalles.plus" -->     
        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>