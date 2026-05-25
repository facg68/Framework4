<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <title>Lista de Presupuestos</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0090"
            SysLockOut


            '
            ' Global  Init()
            '

            dim con, t, l, sqlCommand, sqlString, usu, pre, est, cuantos, sw
            dim suma, mPrecio, tipo, subTotal1, subTotal2, ver, dia, colCampo
            dim aClass, wFecha, wDescripcion, wMontoOrigen, ClaseTabla
            dim wMontoCambio, wContacto, wAplicado, wAcciones
            dim redirectTipo, redirectEstatus, redirectOrdenamiento
            dim labelLocalOriginal, labelLocalCambio

            '
            ' Funciones y Procedimientos
            '

            sub boton(presupuesto, llave, accion, icono, iconoColor, estatus)
                dim iColor, vinculo

                iColor = "azul"
                if iconoColor <> "" then iColor = iconoColor

                if estatus = 0 then
                    '
                    ' El botón aparece deshabilitado
                    '
                    %>
                        <button class="form-btn <%= iColor %> disabled" disabled>
                            <i class=" fa fa-<%= icono %> fa-xl" title="<%= accion %> presupuesto"></i>
                        </button>
                    <%                    
                else
                    vinculo = "pre_det_tra_" & accion & ".asp?m=1&p=" & presupuesto & "&d=" & dia & "&l=" & llave & "&v=" & ver & "&t=" & redirectTipo & "&e=" & redirectEstatus & "&o=" & redirectOrdenamiento
                    %>
                        <button class="form-btn <%= iColor %>" onClick="irA('<%= vinculo %>')">
                            <i class=" fa fa-<%= icono %> fa-xl" title="<%= accion %> presupuesto"></i>
                        </button>
                    <%                        
                end if      
            end Sub

            function NombrePresupuesto(Usuario, Presupuesto)
                dim c, ta

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set ta = c.Execute("SELECT nombre from pre_Presupuesto_Encabezado where (Usuario = '" & usuario & "') AND (Presupuesto = '" & presupuesto & "');")
                        if not (ta.bof or ta.eof) then
                            NombrePresupuesto = ta("nombre")
                        else
                            NombrePresupuesto = ""
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
            end Function

            function SaldoPresupuestoCartera(Usuario, Presupuesto)
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

                            SaldoPresupuestoCartera = db + cr        
                        else
                            SaldoPresupuestoCartera = 0.00
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
            end Function   

            function SaldoPresupuestoEfectivo(Usuario, Presupuesto)
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

                            SaldoPresupuestoEfectivo = db + cr        
                        else
                            SaldoPresupuestoEfectivo = 0.00
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
            end Function           

            function SaldoPresupuestoCarteraTotal(Usuario, Presupuesto)
                dim c, ta, sqlString, db, cr

                sqlString = "SELECT SUM(iif(CuentaOrigen = 'PRE-000', MontoOrigen,0)) AS DB, " & _
                                        " SUM(iif(CuentaDestino = 'PRE-000',MontoDestino,0)) AS CR " & _
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

                            SaldoPresupuestoCarteraTotal = db + cr
                        else
                            SaldoPresupuestoCarteraTotal = 0.00
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
            end Function     

            function SaldoPresupuestoEfectivoTotal(Usuario, Presupuesto)
                dim c, ta, sqlString, db, cr

                sqlString = "SELECT SUM(iif(CuentaOrigen = 'EF-000', MontoOrigen,0)) AS DB, " & _
                                        " SUM(iif(CuentaDestino = 'EF-000',MontoDestino,0)) AS CR " & _
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

                            SaldoPresupuestoEfectivoTotal = db + cr
                        else
                            SaldoPresupuestoEfectivoTotal = 0.00
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
            end Function           

            function TipoPresupuesto(Usuario, Presupuesto)
                dim c, ta

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set ta = c.Execute("SELECT tipo from pre_Presupuesto_Encabezado where (Usuario = '" & usuario & "') AND (Presupuesto = '" & presupuesto & "');")
                        if not (ta.bof or ta.eof) then
                            TipoPresupuesto = ta("tipo")
                        else
                            TipoPresupuesto = ""
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
            end Function      

            function EstatusPresupuesto(Presupuesto)
                dim c, ta

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set ta = c.Execute("SELECT estatus from pre_Presupuesto_Encabezado where (Usuario = '" & Request.Cookies("usuario") & "') AND (Presupuesto = '" & presupuesto & "');")
                        if not (ta.bof or ta.eof) then
                            EstatusPresupuesto = ta("estatus")
                        else
                            EstatusPresupuesto = 0
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
            end Function  

            Function MultiPrecio(Usuario, Presupuesto)
                dim c, ta

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set ta = c.Execute("SELECT multiprecio from pre_Presupuesto_Encabezado where (Usuario = '" & usuario & "') AND (Presupuesto = '" & presupuesto & "');")
                        if not (ta.bof or ta.eof) then
                            MultiPrecio = ta("multiprecio")
                        else
                            MultiPrecio = 0
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
            end function   

            Function Cuantificable(Usuario, Presupuesto)
                dim c, ta

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set ta = c.Execute("SELECT Cuantificable from pre_Presupuesto_Encabezado where (Usuario = '" & usuario & "') AND (Presupuesto = '" & presupuesto & "');")
                        if not (ta.bof or ta.eof) then
                            Cuantificable = ta("Cuantificable")
                        else
                            Cuantificable = 0
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
            end function        

            function FechaHora(Fecha, Hora)
                Dim D, M, A, H, Mi, hCadena

                D = RIGHT("00" & Day(Fecha), 2)
                M = RIGHT("00" & Month(Fecha), 2)
                A = Year(Fecha)

                hCadena = RIGHT("00" & Hora, 4)

                H = LEFT(hCadena, 2)
                Mi = RIGHT(hCadena, 2)

                FechaHora = D & "/" & M & "/" & A & "<br/>" & H & ":" & Mi
            end function

            function LocalMontoOriginal(Presupuesto)
                dim c, ta

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set ta = c.Execute("SELECT MonedaOrigen from pre_Presupuesto_Encabezado where (Usuario = '" & Request.Cookies("Usuario") & "') AND (Presupuesto = '" & presupuesto & "');")
                        if not (ta.bof or ta.eof) then
                            LocalMontoOriginal = ta("MonedaOrigen")
                        else
                            LocalMontoOriginal = ""
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
            end function

            function LocalMontoCambio(Presupuesto)
                dim c, ta

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set ta = c.Execute("SELECT MonedaDestino from pre_Presupuesto_Encabezado where (Usuario = '" & Request.Cookies("Usuario") & "') AND (Presupuesto = '" & presupuesto & "');")
                        if not (ta.bof or ta.eof) then
                            LocalMontoCambio = ta("MonedaDestino")
                        else
                            LocalMontoCambio = ""
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
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

            function CuantasCuentas()
                dim c, tt

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set tt = c.Execute("SELECT COUNT(*) AS Cuantas " & _
                                        "FROM pre_Cuentas " & _
                                        "WHERE Usuario = '" & Request.Cookies("Usuario") & "'" & _
                                        "AND (TipoCuenta = 'A') " & _
                                        "AND (Grupo = 'A');")
                        if (tt.bof or tt.eof) then
                            CuantasCuentas = 0
                        else
                            CuantasCuentas = tt("Cuantas")
                        end if
                    tt.close: set tt = nothing
                c.close: set c = nothing   
            end function      

            Sub Preparar()
                usu = Request.Cookies("usuario")
                pre = Request.QueryString("p")
                redirectTipo = Request.QueryString("t")
                redirectEstatus = Request.QueryString("e")
                redirectOrdenamiento = Request.QueryString("o")
                est = EstatusPresupuesto(pre)
                ver = request.QueryString("v")
                dia = request.QueryString("d")

                colCampo = FondoCampo(redirectTipo, redirectEstatus)
                ClaseTabla = "tabla-" & colCampo

                if est = 0 then 
                    ver = "*"
                else
                    if ver = "" then 
                        ver = "0"         
                    end if
                end if

                if dia = "" then dia = "*"

                suma = Cuantificable(usu, pre)
                mPrecio = MultiPrecio(usu, pre)
                tipo = TipoPresupuesto(usu, pre)
                cuantos = 0

                if mPrecio = 1 then
                    wFecha = 10
                    wDescripcion = 40
                    wMontoOrigen = 10
                    wMontoCambio = 10
                    wAplicado = 5
                    wAcciones = 25
                    labelLocalOriginal = LocalMontoOriginal(pre)
                    labelLocalCambio = LocalMontoCambio(pre)
                else
                    wFecha = 10
                    wDescripcion = 50
                    wMontoOrigen = 0
                    wMontoCambio = 10
                    wAplicado = 5
                    wAcciones = 25
                    labelLocalOriginal = LocalMontoOriginal(pre)          
                end if

                sqlCommand = "SELECT p.Llave, p.Usuario, p.Presupuesto, p.Fecha, p.Hora, p.CuentaOrigen, p.MontoOrigen, p.Descripcion, p.Nota, " & _
                                " p.CuentaDestino, p.MontoDestino, p.MontoCambio, p.Aplicado, c.Nombre AS Contacto, e.MultiPrecio, e.Tipo, e.Estatus as EncEstatus, " & _
                                " dbo.Cripto_Simbolo(e.MonedaOrigen) AS SimboloOrigen, dbo.Cripto_Simbolo(e.MonedaDestino) AS SimboloDestino " & _
                            "FROM dbo.pre_Presupuesto_Detalles AS p " & _
                        "INNER JOIN dbo.pre_Presupuesto_Encabezado AS e " & _
                                "ON p.Presupuesto = e.Presupuesto " & _
                            "AND p.Usuario = e.Usuario " & _
                "LEFT OUTER JOIN (SELECT Codigo, PrimerNombre + ' ' + ISNULL(PrimerApellido, '') AS Nombre " & _
                                    "FROM dbo.con_Contactos AS x " & _
                                    "WHERE (Usuario = '" & usu & "')" & _
                                ") AS c " & _
                                "ON p.Contacto = c.Codigo " & _
                            "WHERE (p.Usuario = '" & usu & "') " & _
                            "AND (p.Presupuesto = '" & pre & "') " 

                    if ver = "1" then sqlCommand = sqlCommand & "AND (p.Aplicado = 1) "
                    if ver = "0" then sqlCommand = sqlCommand & "AND (p.Aplicado = 0) "
                    if dia <> "*" then sqlCommand = sqlCommand & "AND (p.Fecha = '" & dia & "') "

                sqlCommand = sqlCommand & "ORDER BY p.Fecha, p.Hora;"                     
            End Sub    

            function HtmlResumenCuentas()
                dim c, t, html, sqlString

                html = "<table class='tabla tabla-blue' style='width:100%;'><thead>" & _
                        "<tr><th style='text-align:left;'>Cuenta</th><th style='text-align:right;'>Saldo</th></tr>" & _
                        "</thead><tbody>" & _
                        "<tr><td>Cartera</td><td style='text-align:right;'>" & formatNumber(SaldoPresupuestoCartera(usu, pre)) & "</td></tr>" & _
                        "<tr><td>Cartera Presupuestada</td><td style='text-align:right;'>" & formatNumber(SaldoPresupuestoCarteraTotal(usu, pre)) & "</td></tr>" & _
                        "<tr><td>Efectivo</td><td style='text-align:right;'>" & formatNumber(SaldoPresupuestoEfectivo(usu, pre)) & "</td></tr>" & _
                        "<tr><td>Efectivo Presupuestado</td><td style='text-align:right;'>" & formatNumber(SaldoPresupuestoEfectivoTotal(usu, pre)) & "</td></tr>" & _
                        "</tbody></table>"

                HtmlResumenCuentas = html
            end function    

            function FondoCampo(tipo, estatus)
                if tipo = "P" then
                    if estatus = "0" then
                        FondoCampo = "blue"
                    else
                        FondoCampo = "green"
                    end if
                else
                    FondoCampo = "violet"
                end if
            end function                             
        %>

        <style>
            .campo {
                padding: 0.3rem 0.4rem;
                border: 1px solid #ccc;
                border-radius: 0.3rem;
                font-family: 'Ruda', sans-serif;
                font-size: 1rem;
                color: rgb(25, 25, 25);                
                background-color: var(--field-background-color);
                box-sizing: border-box;
                resize: vertical;
            }

            .campo2 {
                padding: 0.3rem 0.4rem;
                border: none;
                font-family: 'Ruda Bold', sans-serif;
                font-size: 1rem;
                color: rgb(25, 25, 25);                
                background-color: transparent !important;
                box-sizing: border-box;
                resize: vertical;
            }   
            
            td {
                font-family: 'Ruda';
                font-size: 16px;
            }

            .swal-title-ruda {
                font-family: 'Arial';
                font-size: 1.5rem;
                font-weight: 600;
                color: #2c2c2c;
            }
            
            /* Ajustes para el <td> con ícono + texto */

                /* 
                    01. Reseteamos el espaciado del <td> para neutralizar
                        estilos impuestos por el framework.

                        Elementos centrados verticalmente 
                        dentro del table-cell (<td>)
                */

                .td-reset {
                    padding: 0              !important;
                    line-height:    normal  !important;
                    vertical-align: middle  !important;
                } 
                
                /* 
                    02. La clase td-flex crea un contenedor
                        del tipo flex con elementos centrados

                        Estos pueden definir nuevas reglas dentro de 
                        su "caja", aunque el <td> (como table-cell) siga 
                        gobernando su altura y alineación dentro de la fila
                */            
                
                .td-flex {
                    display: flex;
                    align-items: center;
                }

                /* 
                    03. Las clases icono-td y texto-td
                        definen las reglas que se aplican
                        dentro del contenedor tipo flex
                        creado con td-flex, por lo que
                        estas reglas son locales al 
                        nuevo contenedor
                */            
            
                .icono-td {
                    display: inline-flex;
                    align-items: center;
                }

                .texto-td {
                    display: inline-flex;
                    align-items: center;
                }

            /* FIN - Ajustes para el <td> con ícono + texto */

            /* Tipos de Transacciones */
                
                .cartera_CR_0 { color: rgb(6, 99, 16); font-size: 15px; font-weight: normal; font-family: 'Ruda Bold';}
                .cartera_CR_1 { color: rgb(123, 150, 126); font-size: 15px; font-family: 'Ruda';}   

                .cartera_DB_0 { color: rgb(114, 14, 14); font-size: 15px; font-weight: normal; font-family: 'Ruda Bold';}
                .cartera_DB_1 { color: rgb(155, 106, 106); font-size: 15px; font-family: 'Ruda';}
                
                .efectivo_CR_0 { color: rgb(9, 37, 101); font-size: 15px; font-weight: normal; font-family: 'Ruda Bold';}
                .efectivo_CR_1 { color: rgb(97, 112, 146); font-size: 15px; font-family: 'Ruda';}
                
                .efectivo_DB_0 { color: rgb(77, 11, 97); font-size: 15px; font-weight: normal; font-family: 'Ruda Bold';}
                .efectivo_DB_1 { color: rgb(134, 101, 145); font-size: 15px; font-family: 'Ruda';}
                
                .otros_0 { color:rgb(0, 0, 0); font-size: 15px; font-weight: normal; font-family: 'Ruda Bold';}
                .otros_1 { color:rgb(116, 116, 116); font-size: 15px; font-family: 'Ruda';} 

            /* Fin - Tipos de Transacciones */              
        </style>                
    </head>

    <body plantilla="tabla" reserva="175">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <% 
            Preparar 

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")
            set t = con.Execute(sqlCommand)
        %>  

        <br />
    
        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 70%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                <span style="font-size: 15px; text-align: left;">
                    <input name="preCodigo" id="preCodigo" 
                           type="text" value="<%= pre %>" 
                           class="campo2" 
                           style="width: 135px;"
                           onchange="cambiarCodigo()"
                    >
                           
                    &nbsp;

                    <input name="preNombre" id="preNombre" 
                           type="text" value="<%= NombrePresupuesto(usu, pre) %>" 
                           class="campo2" 
                           style="width: 450px;"
                           onchange="cambiarNombre()"
                    >
                </span>

                <br />

                <span>
                    <%
                        sqlString = "SELECT DISTINCT TOP (100) PERCENT Fecha, dbo.Cripto_FechaFormulario(Fecha) AS FF " & _
                                        "FROM dbo.pre_Presupuesto_Detalles AS p " & _
                                    "WHERE (Presupuesto = '" & pre & "') " & _
                                        "AND (Usuario = '" & usu & "') " 

                        if ver <> "*" then sqlString = sqlString & "AND (Aplicado = " & ver & ") "

                        sqlString = sqlString & "ORDER BY Fecha DESC;"

                        set l = con.Execute(sqlString)
                    %>

                        <select class="campo" name="cboDia" id="cboDia" onchange="recargarMostrar();" <% if est = 0 then response.write " disabled" %>>
                            <option value="*" <% if dia = "*" then response.write " selected" %>>- - Todos - -</option>
                            <%
                                if not (l.bof or l.eof) then
                                    do
                                        response.write "<option value='" & l("Fecha") & "' " 
                                            if dia = l("Fecha") then 
                                                response.write " selected"
                                            end if
                                        response.write ">" & left(l("FF") , 10) & "</option>"
                                    
                                        l.MoveNext
                                    loop until l.eof
                                end if
                            %>
                        </select>
                        
                    <% l.close: set l = nothing %>

                    <select class="campo" name="cboVer" id="cboVer" onchange="recargarMostrar();" <% if est = 0 then response.write " disabled" %>>
                        <option value="0" <% if ver = "0" then response.write " selected" %>>Ver Pendientes</option>
                        <option value="1" <% if ver = "1" then response.write " selected" %>>Ver Cerradas</option>     
                        <option value="*" <% if ver = "*" then response.write " selected" %>>Ver Todas</option>                                                 
                    </select>  

                    <select class="campo" name="cboCuantificable" id="cboCuantificable" onchange="aplicarCuantificable();" <% if est = 0 then response.write " disabled" %>>
                        <option value="1" <% if suma = 1 then response.write " selected" %>>Afecta Montos</option>
                        <option value="0" <% if suma = 0 then response.write " selected" %>>Se ignora</option>
                    </select>                 
                </span>
            </div>
            
            <div style="flex: 0 0 30%; text-align: right;">
                <%
                    estatus = ""
                    if est = 0 then estatus = " disabled "
                %>

                &nbsp;&nbsp;               

                <button class="form-btn azul tiny" onClick="irA('../lista.asp')">
                    <i class=" fa fa-backward fa-xl" title="Volver a la Lista de Presupuestos"></i>
                </button>

                <button class="form-btn azul tiny"onClick="imprimir('<%= pre %>')">
                    <i class=" fa fa-print fa-xl" title="Visualización de Impresión"></i>
                </button>     

                <button type="button" class="form-btn tiny violeta" onclick="verResumenCuentas();">
                    <i class="fa fa-wallet fa-xl"></i>
                </button>                                     

                <button class="form-btn tiny naranja <%= estatus %>" onClick="cerrar('<%= pre %>')" <%= estatus %>> 
                    <i class=" fa fa-lock fa-xl" title="Cerrar este presupuesto"></i>
                </button>

                <button class="form-btn tiny rojo<%= estatus %>" onClick="borrar()" <%= estatus %>>
                    <i class=" fa fa-trash fa-xl" title="Borrar este presupuesto"></i>
                </button>
            </div>
        </div>   

        <div class="no-ver">
            <input id="ordenamiento" name="ordenamiento"  value="<%= oParan %>">
            <input id="txtPre" name="txtPre"  value="<%= pre %>">                   

            <input id="req_Tipo"          name="req_Tipo"          value="<%= redirectTipo %>">
            <input id="req_Estatus"       name="req_Estatus"       value="<%= redirectEstatus %>">
            <input id="req_Ordenamiento"  name="req_Ordenamiento"  value="<%= redirectOrdenamiento %>">     
        </div>     

        <div class="main" style="width: 95%;">
            <div class="line">
                <div class="tabla-wrapper">
                    <table class="tabla <%= ClaseTabla %>">
                        <thead>
                            <% if mPrecio = 1 then %>
                                <th class="sticky"  style="text-align: center; width:<%= wAplicado    %>%;" >&nbsp;</th>
                                <th class="sticky"  style="text-align: center; width:<%= wFecha       %>%;" >Fecha</th>
                                <th class="sticky"  style="text-align: center; width:<%= wDescripcion %>%;" >Descripcion</th>
                                <th class="sticky"  style="text-align: center; width:<%= wMontoOrigen %>%;" >Origen (<%= labelLocalOriginal %>)</th>
                                <th class="sticky"  style="text-align: center; width:<%= wMontoCambio %>%;" >Destino (<%= labelLocalCambio %>)</th>
                                <th class="sticky"  style="text-align: center; width:<%= wAcciones    %>%;" >Acciones</th>
                            <% else %>  
                                <th class="sticky"  style="text-align: center; width:<%= wAplicado    %>%;" >&nbsp;</th>
                                <th class="sticky"  style="text-align: center; width:<%= wFecha       %>%;" >Fecha</th>
                                <th class="sticky"  style="text-align: center; width:<%= wDescripcion %>%;" >Descripcion</th>
                                <th class="sticky"  style="text-align: center; width:<%= wMontoCambio %>%;" >Monto</th>
                                <th class="sticky"  style="text-align: center; width:<%= wAcciones    %>%;" >Acciones</th> 
                            <% end if %>
                        </thead>

                        <tbody>
                            <%
                                if not (t.bof or t.eof) then
                                    Do
                                        cuantos = cuantos + 1

                                        if (est = 0) then
                                            '
                                            ' El presupuesto está cerrado, pero se ha abierto para consulta...
                                            ' Presentamos los detalles con colores vivos (tipo "abierto")
                                            ' para ver mejor los datos. Igual no puede ser modificado.
                                            '
                                            aClass = claseVinculo(t("CuentaOrigen"), t("CuentaDestino"), t("MontoDestino"), 0)
                                        else
                                            aClass = claseVinculo(t("CuentaOrigen"), t("CuentaDestino"), t("MontoDestino"), t("Aplicado"))
                                        end if

                                        vinculo = "pre_det_tra_editar.asp?p=" & pre & "&d=" & dia & "&l=" & t("Llave") & "&v=" & ver & "&t=" & redirectTipo & "&e=" & redirectEstatus & "&o=" & redirectOrdenamiento

                                        %>
                                            <tr>
                                                <td style="text-align: center;">
                                                    <%
                                                        if tipo = "P" then
                                                            if t("aplicado") = 0 then
                                                                chk_vinculo = "pre_det_tra_aplicar.asp?p=" & pre & "&d=" & dia & "&l=" & t("Llave") & "&v=" & ver & "&t=" & redirectTipo & "&e=" & redirectEstatus & "&o=" & redirectOrdenamiento                      

                                                                %><img src="imagenes/unchecked.png" style="border: none;" onclick="irA('<%= chk_vinculo %>')"><%
                                                            else
                                                                response.write "<img src='imagenes/checked.png' border='0'>"
                                                            end if
                                                        else
                                                            response.write "&nbsp;"
                                                        end if                                                                
                                                    %>
                                                </td>

                                                <td class="<%= aClass %>" style="text-align: center;" onClick="irA('<%= vinculo %>')">
                                                    <%= FechaHora(t("Fecha"),  t("Hora"))  %>
                                                </td>

                                                <td class="<%= aClass %> td-reset" style="text-align: left;" onClick="irA('<%= vinculo %>')">
                                                    <div class="td-flex">
                                                        <span class="icono-td">
                                                            <% if (len(trim(t("Nota"))) > 0 ) then %>
                                                                &nbsp;&nbsp;&nbsp;<img src="imagenes/nota.png" width="23" height="23">&nbsp;&nbsp;
                                                            <% else %>
                                                                &nbsp;&nbsp;&nbsp;
                                                            <% end if %>
                                                        </span>
                                                        <span class="texto-td"><%= t("Descripcion") %></span>
                                                    </div>
                                                </td>

                                                <% if mPrecio = 1 then %>
                                                    <td class="<%= aClass %>" style="text-align: right;" onClick="irA('<%= vinculo %>')">
                                                        <%= FormatNumber(t("MontoDestino")) %>
                                                    </td>
                                                <% end if %>

                                                <td class="<%= aClass %>" style="text-align: right;" onClick="irA('<%= vinculo %>')">
                                                    <%= FormatNumber(t("MontoCambio")) %>
                                                </td>

                                                <td style="text-align: center;">
                                                    <%
                                                        if t("tipo") = "P" then
                                                            if  t("encEstatus") = 0 then
                                                                boton pre, t("Llave"), "duplicar", "copy", "azul", 0
                                                                boton pre, t("Llave"), "borrar", "trash", "azul", 0
                                                            else
                                                                boton pre, t("Llave"), "nota", "edit", "verde", 1
                                                                boton pre, t("Llave"), "duplicar", "copy", "verde", 1

                                                                if t("aplicado") = 0 then
                                                                    boton pre, t("Llave"), "borrar", "trash", "verde", 1
                                                                else
                                                                    boton pre, t("Llave"), "borrar", "trash", "verde", 0
                                                                end if
                                                            end if
                                                        else
                                                            boton pre, t("Llave"), "nota", "edit", "violeta", 1
                                                            boton pre, t("Llave"), "duplicar", "copy", "violeta", 1
                                                            boton pre, t("Llave"), "borrar", "trash", "violeta", 1
                                                        end if                                                                                                                                    
                                                    %>
                                                </td>
                                            </tr>                                                                                                                                                                                                                               
                                        <%

                                        t.MoveNext
                                    Loop Until t.eof                                         
                                end if
                            %>                                                       
                        </tbody>                        

                        <tfoot>
                            <tr>
                                <%
                                    tCols = 5
                                    if mPrecio = 1 then tCols = 6
                                %>
                                <td colspan="<%= (tCols - 2) %>" class="sticky"  style="text-align: center; font-weight: normal;">
                                    <%
                                        Select Case cuantos
                                            case 0: response.write "No se encontraron Transaccioness"
                                            case 1: response.write "Sólo se encontró una Transacción"
                                            case else
                                                response.write "Se encontraron " & Cuantos &  " Transacciones"
                                        end Select
                                    %>                                            
                                </td>

                                <td colspan="2" class="sticky"  style="text-align: right;" style="text-align: center;">
                                    <%
                                        if est <> 0 then            
                                            vinculo = "pre_det_tra_editar.asp?p=" & pre & "&d=" & dia & "&l=-1&v=" & ver & "&t=" & redirectTipo & "&e=" & redirectEstatus & "&o=" & redirectOrdenamiento

                                            if mPrecio = 1 then
                                                %>
                                                    <button type="button" 
                                                            class="form-btn naranja" 
                                                            style="width: 125px; font-size: 16px; color: white;"
                                                            onClick="irA('<%= "pre_det_recalcular.asp?m=1&p=" & pre %>')">
                                                        Recalcular
                                                    </button>
                                                <% 
                                            end if %>

                                                &nbsp;

                                                <button type="button" 
                                                        class="form-btn azul" 
                                                        style="width: 125px; font-size: 16px; color: white;"
                                                        onClick="irA('<%= vinculo %>')">
                                                    Transaccion
                                                </button>
                                            <% 
                                        else 
                                            response.write "&nbsp;"
                                        end if
                                    %>                                            
                                </td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
        </div>
        
        <br /><br />   

        <% t.close: set t = nothing %>        

        <script type="text/javascript">
            function irA(vinculo) {
                window.location.href = vinculo;
            }

            function recargarMostrar() {
                var d = document.getElementById("cboDia").value;
                var v = document.getElementById("cboVer").value;
                var p = document.getElementById("txtPre").value;
                var t = document.getElementById("req_Tipo").value;
                var e = document.getElementById("req_Estatus").value;
                var o = document.getElementById("req_Ordenamiento").value;

                var pagina = "pre_det_editar.asp?p=" + p + "&d=" + d + "&v=" + v + "&t=" + t + "&e=" + e + "&o=" + o;

                window.location.href = pagina;
            }

            function borrar() {
                var confirmacion = confirm("Desea borrar el presupuesto <%= pre %>?");
                var vinculo = "pre_det_borrar.asp?p=<%= pre %>&t=<%= redirectTipo %>&e=<%= redirectEstatus %>&o=<%= redirectOrdenamiento %>";

                if (confirmacion) {
                    confirmacion = confirm("Esta opcion AFECTA los totales de las Cuentas si no se ha hecho un cierre. ¿Esta COMPLETAMENTE SEGURO?");            

                    if (confirmacion) {
                        confirmacion = confirm("Esta es la ultima advertencia. ¿Quiere Borrar Este Presupuesto?");       

                        if (confirmacion) {     
                            window.location.href = vinculo;
                        };
                    };
                };
            }    

            function cerrar(codigo) {
                var confirmacion = confirm("Esta seguro de cerrar el presupuesto " + codigo + "?");
                var vinculo = "pre_det_cerrar.asp?p=" + codigo

                if (confirmacion) {
                    confirmacion = confirm("Esta opcion BLOQUEA este presupuesto para siempre. ¿Esta COMPLETAMENTE SEGURO?");            
                    if (confirmacion) {
                        window.location.href = vinculo;
                    }
                }        
            }  

            function imprimir(codigo) {
                var vinculo = "pre_det_imprimir.asp?p=" + codigo;
                window.location.href = vinculo;
            }     

            function cambiarCodigo() {
                var confirmacion = confirm("Desea Cambiar El Código de Este Presupuesto?");

                if (confirmacion) {
                    var codigo = document.getElementById("preCodigo").value;
                    var pre = document.getElementById("txtPre").value;

                    var vinculo = "pre_det_afectar_codigo.asp?c=" + codigo + "&p=" + pre;
                    navigator.sendBeacon(vinculo);

                    Swal.fire({
                        icon: 'success',
                        title: 'Actualizado',
                        text: 'El Código fue modificado correctamente',
                        timer: 1500,
                        showConfirmButton: false,
                        position: 'top-end',
                        toast: true,
                        customClass: {
                            title: 'swal-title-ruda'
                        }
                    });
                }
            }   
            
            function cambiarNombre() {
                var confirmacion = confirm("Desea Cambiar El Nombre de Este Presupuesto?");

                if (confirmacion) {
                    var nombre = document.getElementById("preNombre").value;
                    var pre = document.getElementById("txtPre").value;
                    
                    var vinculo = "pre_det_afectar_nombre.asp?p=" + encodeURIComponent(pre) + "&n=" + encodeURIComponent(nombre);
                    navigator.sendBeacon(vinculo);

                    Swal.fire({
                        icon: 'success',
                        title: 'Actualizado',
                        text: 'El Nombre fue modificado correctamente',
                        timer: 1500,
                        showConfirmButton: false,
                        position: 'top-end',
                        toast: true,
                        customClass: {
                            title: 'swal-title-ruda'
                        }
                    });   
                }            
            }     

            function aplicarCuantificable() {
                var confirmacion = confirm("Desea Cambiar El Tipo de Presupuesto?");

                if (confirmacion) {
                    var pre = document.getElementById("txtPre").value;
                    var suma = document.getElementById("cboCuantificable").value;

                    var vinculo = "pre_det_afectar_cuantificable.asp?p=" + pre + "&s=" + suma;
                    navigator.sendBeacon(vinculo);                    

                    Swal.fire({
                        icon: 'success',
                        title: 'Actualizado',
                        text: 'El Tipo de Presupuesto fue modificado correctamente',
                        timer: 1500,
                        showConfirmButton: false,
                        position: 'top-end',
                        toast: true,
                        customClass: {
                            title: 'swal-title-ruda'
                        }
                    });   
                }            
            }  

            function verResumenCuentas() {
                Swal.fire({
                    title: 'Resumen de Presupuesto',
                    html: `<%= Replace(HtmlResumenCuentas(), "`", "\`") %>`,
                    width: 600,
                    showConfirmButton: true,
                    confirmButtonText: 'Cerrar',
                    confirmButtonColor: '#2f6fb2',
                    customClass: {
                        popup: 'swal2-border-radius',
                        title: 'swal-title-ruda'
                    }
                });
            }   
        </script>

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>