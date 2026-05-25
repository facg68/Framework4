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
            ' Init()
            '
            dim con, t, sqlString, vinculo, sw, tipoPre, estatusPre, cuantos
            dim o1, o2, o3, o4, o5, oParam, orden, colCampo, ClaseTabla

            tipoPre = request.Cookies("listaPreTipo")
            estatusPre = request.Cookies("listaPreEstatus")
            oParam = request.Cookies("listaPreOrdenamiento")

            if (isnull(tipoPre) or tipoPre = "")        then tipoPre = "P"
            if (isnull(estatusPre) or estatusPre = "")  then estatusPre = "1"
            if (isnull(oParam) or oParam = "")          then oParam = "PA"

            ActualizarTodasLasCuentas
            ActualizarTotales
            ActualizarFechas
            
            cuantos = 0
            ProcesarOrdenamiento oParam
            FondoCampo tipoPre, estatusPre

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
                        "ORDER BY " & orden & ";"

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")    
            set t = con.Execute(sqlString)

            '
            ' Funciones y Procedimientos
            '

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

            sub boton(usuario, presupuesto, accion, icono, iconoColor, estatus)
                dim iColor, vinculo

                iColor = "naranja"
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
                    select case accion
                        case "convertir"
                            %>
                                <button class="form-btn <%= iColor %>" onClick="modelo('<%= presupuesto %>')">
                                    <i class=" fa fa-<%= icono %> fa-xl" title="<%= accion %> presupuesto"></i>
                                </button>
                            <%
                        case else
                            vinculo = "presupuestos/pre_det_" & accion & ".asp?p=" & Presupuesto

                            %>
                                <button class="form-btn <%= iColor %>" onClick="irA('<%= vinculo %>')">
                                    <i class=" fa fa-<%= icono %> fa-xl" title="<%= accion %> presupuesto"></i>
                                </button>
                            <%       
                    end select
                end if
            end Sub

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

            sub ProcesarOrdenamiento(llave)
                o1 = "PA"
                o2 = "NA"
                o3 = "DA"
                o4 = "HA"
                o5 = "SA"

                select case llave
                    case "PD" 
                        o1 = "PA"
                        orden = "Presupuesto DESC"

                    case "ND" 
                        o2 = "NA"
                        orden = "Nombre DESC"

                    case "DD" 
                        o3 = "DA"
                        orden = "Desde DESC"

                    case "HD" 
                        o4 = "HA"
                        orden = "Hasta DESC"

                    case "SD" 
                        o5 = "SA"
                        orden = "Saldo DESC"

                    case "PA" 
                        o1 = "PD"
                        orden = "Presupuesto ASC"

                    case "NA" 
                        o2 = "ND"
                        orden = "Nombre ASC"

                    case "DA" 
                        o3 = "DD"
                        orden = "Desde ASC"

                    case "HA" 
                        o4 = "HD"
                        orden = "Hasta ASC"

                    case "SA" 
                        o5 = "SD"
                        orden = "Saldo ASC"
                end select      
            end sub

            function HtmlResumenCuentas()
                dim c, t, html, sqlString

                sqlString = "SELECT Codigo, Nombre, dbo.pre_SaldoCuenta(Usuario, Codigo, LocalMonetario) AS SaldoActual " & _
                            "FROM pre_Cuentas " & _
                            "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') AND (TipoCuenta = 'A') AND (Grupo = 'A') " & _
                            "ORDER BY Nombre;"    

                html = "<table class='tabla tabla-blue' style='width:100%;'>" & _
                            "<thead>" & _
                                "<tr>" & _
                                    "<th style='text-align:left;'>Cuenta</th>" & _
                                    "<th style='text-align:right;'>Saldo</th>" & _
                                "</tr>" & _
                            "</thead><tbody>"

                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")

                    set t = c.Execute(sqlString)
                    do while not t.eof
                        val = "0.00"
                        if isnumeric(t("SaldoActual")) then val = FormatNumber(t("SaldoActual"), 2)

                        html = html & "<tr>" & _
                                        "<td>" & t("Nombre") & "</td>" & _
                                        "<td style='text-align:right;'>" & val & "</td>" & _
                                    "</tr>"
                        t.movenext
                    loop

                    t.close : set t = nothing
                c.close : set c = nothing

                html = html & "</tbody></table>"

                HtmlResumenCuentas = html
            end function 

            Sub FondoCampo(tipo, estatus)
                if tipo = "P" then
                    if estatus = 0 then
                        colCampo = "#f0faff"
                        ClaseTabla = "tabla-blue"
                    else
                        colCampo = "#F3FAEE"
                        ClaseTabla = "tabla-green"
                    end if
                else
                    colCampo = "#f5eefaff"
                    ClaseTabla = "tabla-violet"
                end if
            End Sub
        %>

        <style>
            .campo {
                padding: 0.3rem 0.4rem;
                border: 1px solid #ccc;
                border-radius: 0.3rem;
                font-family: 'Ruda', sans-serif;
                font-size: 1rem;
                color: rgb(25, 25, 25);                
                background-color: <%= colCampo %>;
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
        </style>                
    </head>

    <body plantilla="tabla" tabla="100" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <br />
    
        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                Presupuestos
            </div>
            
            <div style="flex: 0 0 50%; text-align: right;">
                <form name="filtro" id="filtro" action="preCookies.asp" method="post">
                    <input class="no-ver" id="ordenamiento" name="ordenamiento" value="<%= oParam %>">

                    Ver&nbsp;
                    <select class="campo" name="cboTipoPresupuesto" id="cboTipoPresupuesto" onChange="Requery();">
                        <option value="P" <% if tipoPre = "P" then response.write " selected" %>>Presupuestos</option>
                        <option value="M" <% if tipoPre = "M" then response.write " selected" %>>Modelos</option>
                    </select>    

                    &nbsp;&nbsp;               

                    <select class="campo" name="cboEstatusPresupuesto" id="cboEstatusPresupuesto" onChange="Requery();">
                        <option value="1" <% if estatusPre = "1" then response.write " selected" %>>Abiertos</option>
                            <% if tipoPre = "P" then %>
                                <option value="0" <% if estatusPre = "0" then response.write " selected" %>>Cerrados</option>
                            <% end if %>
                        <option value="2" <% if estatusPre = "2" then response.write " selected" %>>Archivados</option>
                    </select>     

                    &nbsp;&nbsp;     

                    <button type="button" class="form-btn tiny violeta" onclick="verResumenCuentas();">
                        <i class="fa fa-wallet fa-xl"></i>
                    </button>                    

                    &nbsp;

                    <button type="button" class="form-btn tiny azul" onClick="irA('presupuestos/pre_enc_nuevo.asp')">
                        <i class="fa fa-plus fa-xl"></i>
                    </button>
                </form>
            </div>
        </div>        

        <div class="main" style="width: 95%;">
            <div class="line">
                <div class="tabla-wrapper">
                    <table class="tabla <%= ClaseTabla %>">
                        <thead>
                            <tr>
                                <th class="sticky" style="width:15%; text-align: center;" onClick="ordenar('<%= o1 %>');">Presupuesto</th>
                                <th class="sticky" style="width:40%; text-align: center;" onClick="ordenar('<%= o2 %>');">Nombre</th>
                                <th class="sticky" style="width:15%; text-align: center;" onClick="ordenar('<%= o5 %>');">Saldo</th>
                                <th class="sticky" style="width:30%; text-align: center;">Acciones</th>
                            </tr>
                        </thead>

                        <tbody>
                            <%
                                if not (t.bof or t.eof) then
                                    Do
                                        cuantos = cuantos + 1
                                        vinculo = "presupuestos/pre_det_editar.asp?p=" & t("Presupuesto") & "&t=" & tipoPre & "&e=" & estatusPre & "&o=" & oParam

                                        %>
                                            <tr>
                                                <td style="text-align: center;" onClick="irA('<%= vinculo %>')">
                                                    <%= t("Presupuesto") %>
                                                </td>

                                                <td style="text-align: left;" onClick="irA('<%= vinculo %>')">
                                                    <%= t("Nombre") %><br/>
                                                    ( <%= FechaForm(t("Desde")) %> - <%= FechaForm(t("Hasta")) %> )
                                                </td>

                                                <td style="text-align: right;" onClick="irA('<%= vinculo %>')">
                                                    <%= FormatNumber(t("Saldo")) %>
                                                </td> 

                                                <td style="text-align: center;">
                                                    <%
                                                        if t("tipo") = "P" then
                                                            select case t("estatus")
                                                                case 0
                                                                    boton t("usuario"), t("presupuesto"), "imprimir",   "print",    "azul", 1

                                                                case 1
                                                                    boton t("usuario"), t("presupuesto"), "copiar",     "copy",     "verde", 1
                                                                    boton t("usuario"), t("presupuesto"), "convertir",  "gear",     "verde", 1
                                                                    boton t("usuario"), t("presupuesto"), "archivar",   "archive",  "verde", 1
                                                                    boton t("usuario"), t("presupuesto"), "imprimir",   "print",    "verde", 1

                                                                case 2
                                                                    boton t("usuario"), t("presupuesto"), "copiar",     "copy",     "verde", 1
                                                                    boton t("usuario"), t("presupuesto"), "convertir",  "gear",     "verde", 1
                                                                    boton t("usuario"), t("presupuesto"), "archivar",   "archive",  "verde", 1
                                                                    boton t("usuario"), t("presupuesto"), "imprimir",   "print",    "verde", 1                              
                                                            end select
                                                        else
                                                            '
                                                            ' Los modelos no tienen estatus "cerrado" (0)
                                                            '
                                                            boton t("usuario"), t("presupuesto"), "editar",   "edit",    "violeta", 1
                                                            boton t("usuario"), t("presupuesto"), "copiar",   "copy",    "violeta", 1
                                                            boton t("usuario"), t("presupuesto"), "archivar", "archive", "violeta", 1
                                                            boton t("usuario"), t("presupuesto"), "imprimir", "print",   "violeta", 1
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
                                <td colspan="4" class="sticky" style="text-align: center; font-weight: normal;">
                                    <%
                                        nObjeto = "Modelo"
                                        if tipoPre = "P" then nObjeto = "Presupuesto"

                                        Select Case cuantos
                                            case 0: response.write "No se encontraron " & nObjeto & "s"
                                            case 1: response.write "Sólo se encontró un " & nObjeto
                                            case else
                                                response.write "Se encontraron " & Cuantos &  " " & nObjeto & "s"
                                        end Select
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
            function Requery() {
                document.getElementById("filtro").submit();
            }

            function ordenar(campo) {
                document.getElementById("ordenamiento").value = campo;  
                document.getElementById("filtro").submit();
            }

            function modelo(presupuesto) {
                var alerta = "Está seguro de convertir el presupuesto " + presupuesto + " en un modelo?";
                var vinculo = "presupuestos/pre_det_convertir.asp?p=" + presupuesto;

                if (confirm(alerta)) {
                    window.location.href = vinculo;
                } else {
                    alert("Proceso Cancelado.");
                }     
            }

            function borrar(codigo) {
                var confirmacion = confirm("Esta seguro de borrar el presupuesto " + codigo + "?");
                var vinculo = "presupuestos/pre_det_borrar.asp?p=" + codigo

                if (confirmacion) {
                    window.location.href = vinculo;
                } else {
                    alert("Proceso Cancelado.");        
                }        
            }       

            function irA(vinculo) {
                window.location.href = vinculo;
            }

            function verResumenCuentas() {
                Swal.fire({
                    title: 'Resumen de cuentas',
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