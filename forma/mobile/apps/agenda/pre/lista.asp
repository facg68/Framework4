<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Presupuestos" %>
        <title><%= PageTitle %></title>

        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0090"
            SysLockOut

            ' Init --------------------------------------------------------------------------------------- 
                dim con, t, sqlString, titulo, sw, tipoPre, estatusPre, clasePre, cuantos

                tipoPre = request.Cookies("pre_t")
                estatusPre = request.Cookies("pre_e")

                if (isnull(tipoPre) or tipoPre = "")        then tipoPre = "P"
                if (isnull(estatusPre) or estatusPre = "")  then estatusPre = "1"

                if tipoPre = "P" then
                    select case estatusPre
                        case "0"
                            clasePre = "azul"
                            titulo = "Presupuestos Cerrados"
                        case "1"
                            clasePre = "verde"
                            titulo = "Presupuestos Abiertos"
                        case "2"
                            clasePre = "verde"
                            titulo = "Presupuestos Archivados"
                    end select
                else
                    clasePre = "morado"

                    if estatusPre = "1" then
                        titulo = "Modelos de Presupuesto"
                    else
                        titulo = "Modelos Archivados"
                    end if
                end if

                ActualizarTodasLasCuentas
                ActualizarTotales
                ActualizarFechas
            
                cuantos = 0

                '
                ' Creamos la cadena de conexión, dependiendo de los
                ' datos del filtro, o generamos una cadena nueva

                sqlString = "SELECT Usuario, Presupuesto, Tipo, Nombre, Desde, Hasta," & _
                                " dbo.Cripto_CambiarMoneda(SaldoFinal, MonedaOrigen, dbo.Cripto_UsuLocal('" &  Request.Cookies("usuario")  & "')) AS Saldo," & _
                                " Estatus, MultiPrecio " & _
                            "FROM dbo.pre_Presupuesto_Encabezado AS p " & _
                            "WHERE (Usuario = '" & Request.Cookies("usuario") & "') " & _
                            "AND (Tipo = '" & tipoPre & "') " & _
                            "AND (Estatus = " & estatusPre & ") " & _
                            "ORDER BY Presupuesto ASC;"

                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")    
                set t = con.Execute(sqlString)      
            '----------------------------------------------------------------------------------------------

            ' Funciones y Procedimientos ------------------------------------------------------------------
                sub ActualizarTodasLasCuentas()
                    dim c 

                    set c = server.CreateObject("ADODB.Connection")
                    c.open Application("Conn")
                        c.Execute("exec pre_Enc_ActualizarSaldoCuentas '" & Request.Cookies("Usuario") & "'")
                    c.close: set c = nothing
                end sub

                sub ActualizarTotales()
                    dim c 

                    set c = server.CreateObject("ADODB.Connection")
                    c.open Application("Conn")
                        c.Execute("exec pre_Enc_ActualizarSaldoListas '" & Request.Cookies("Usuario") & "'")
                    c.close: set c = nothing
                end sub

                sub ActualizarFechas()
                    dim c 

                    set c = server.CreateObject("ADODB.Connection")
                    c.open Application("Conn")
                        c.Execute("exec pre_Enc_ActualizarFechas '" & Request.Cookies("Usuario") & "'")
                    c.close: set c = nothing
                end sub      

                Function FechaForm(FechaDB)
                    dim a, m, d

                    FechaForm = ""

                    if not isnull(FechaDB) then
                        d = RIGHT("00" & day(FechaDB) ,2)
                        m = RIGHT("00" & month(FechaDB), 2)
                        A = year(FechaDB)

                        FechaForm = d & "/" & m & "/" & a
                    end if
                end function    

                function HtmlResumenCuentas()
                    dim c, t, html, sqlString

                    sqlString = "SELECT Codigo, Nombre, dbo.pre_SaldoCuenta(Usuario, Codigo, LocalMonetario) AS SaldoActual " & _
                                "FROM pre_Cuentas " & _
                                "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') AND (TipoCuenta = 'A') AND (Grupo = 'A') " & _
                                "ORDER BY Nombre;"

                    html = "<div class='cuentas-resumen'>"

                    set c = Server.CreateObject("ADODB.Connection")
                    c.open Application("Conn")
                        set t = c.Execute(sqlString)
                            do while not t.eof
                                html = html & _
                                    "<div class='cuenta-row'>" & _
                                        "<div class='cuenta-nombre'>" & t("Nombre") & "</div>" & _
                                        "<div class='cuenta-saldo-linea'>" & _
                                            "<span class='saldo-label'>Saldo Actual:</span> " & _
                                            "<span class='cuenta-saldo'>" & FormatNumber(t("SaldoActual"), 2) & "</span>" & _
                                        "</div>" & _
                                    "</div>"

                                t.movenext
                            loop
                        t.close : set t = nothing
                    c.close : set c = nothing

                    html = html & "</div>"
                    HtmlResumenCuentas = html
                end function    

                function TransaccionesAplicadas(Presupuesto)
                    dim c, t, html, sqlString

                    sqlString = "SELECT Presupuesto, Usuario, ISNULL(COUNT(*), 0) AS ta " & _
                                "FROM dbo.pre_Presupuesto_Detalles " & _
                                "WHERE (Aplicado = 1) " & _
                                "AND (Presupuesto = '" & Presupuesto & "') " & _
                                "AND (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                "GROUP BY Presupuesto, Usuario;"

                    set c = Server.CreateObject("ADODB.Connection")
                    c.open Application("Conn")
                        set t = c.Execute(sqlString)
                            if (t.bof or t.eof) then
                                TransaccionesAplicadas = 0
                            else
                                TransaccionesAplicadas = t("ta")
                            end if
                        t.close : set t = nothing
                    c.close : set c = nothing
                end function
            '----------------------------------------------------------------------------------------------
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
                .pre-morado { background-color:#6b3194; }             
                .presupuesto-item { padding-left: 16px; }            
                .presupuesto-item:active { background: rgba(0,0,0,0.06); }

                .presupuesto-info {
                    flex: 1;
                    min-width: 0;
                }

                .presupuesto-nombre {
                    font-family: 'Ruda Bold';
                    font-size: 1.15rem;
                    color: #111;
                    white-space: normal;        
                    overflow: visible;          
                    text-overflow: clip;        
                    word-break: break-word;     
                }            

                .presupuesto-datos {
                    font-family: 'Ruda Bold';
                    font-size: 1rem;
                    color: #555;
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
        </style>                
    </head>

    <body plus= ".presupuesto-item : plus_telPre; ">
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <div class="page-title-bar pre-<%= clasePre %>">
            <%= titulo %>
        </div>

        <main>
            <%
                if not (t.bof or t.eof) then
                    cuantos = 0

                    Do
                        cuantos = cuantos + 1

                        %>
                            <a class="presupuesto-item" 
                               id="<%= t("Presupuesto") %>"
                               data-presupuesto="<%= t("Presupuesto") %>"
                               data-tipo="<%= tipoPre %>"
                               data-estatus="<%= estatusPre %>"
                               data-transacciones="<%= TransaccionesAplicadas(t("Presupuesto")) %>"
                            >
                                <div class="contacto-info">
                                    <div class="presupuesto-nombre">
                                        <div class="contacto-nombre">
                                            <%= t("Nombre") %>
                                        </div>
                                    </div>

                                    <div class="presupuesto-datos">
                                        <div class="contacto-telefono">
                                            <%
                                                response.write "Fechas: (" & FechaForm(t("Desde")) & " - " & FechaForm(t("Hasta")) & ")<br/>"
                                                response.write "Saldo Final Presupuestado: " & FormatNumber(t("Saldo")) 
                                            %>
                                        </div>
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

            function nuevo() {
                var vinculo = "./presupuestos/pre_enc_nuevo.asp";
                window.location.href = vinculo;                
            }            
        </script>

        <!-- #include virtual = "/forma/plus/telPresupuestos.plus" -->     
        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>