<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Crear Nuevo Prespuesto" %>
        <title><%= PageTitle %></title>

        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0090"
            SysLockOut


            Function Pad2(valor)
                Pad2 = Right("0" & valor, 2)
            End Function

            Function FormatearFecha(fecha)
                FormatearFecha = Pad2(Day(fecha)) & "/" & _
                                Pad2(Month(fecha)) & "/" & _
                                Year(fecha)
            End Function

            Function FechaHoy()
                FechaHoy = FormatearFecha(Date)
            End Function

            Function FechaFinal()
                FechaFinal = FormatearFecha(DateAdd("d", 15, Date))
            End Function

            Function HoraHoy()
                HoraHoy = Pad2(Hour(Time)) & ":" & _
                        Pad2(Minute(Time))
            End Function

            Function MonedaUsuario()
                MonedaUsuario = ParametroUsuario("agenda", "pre_local_monetario")

                if isnull(MonedaUsuario) = True then
                    MonedaUsuario = "US"
                end if
            end Function

            function IncrementoPresupuesto(presupuesto)
                dim c, sqlString, tt 

                sqlString = "select dbo.pre_diasModelo(Usuario, Presupuesto) as dias " & _
                            "from pre_Presupuesto_Encabezado " & _
                            "where presupuesto = '" & presupuesto & "' " & _
                            "and Usuario = '" & request.Cookies("Usuario") & "';"

                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set tt = c.execute(sqlString)
                        if (tt.bof or tt.eof) then
                            IncrementoPresupuesto = 0
                        else
                            IncrementoPresupuesto = tt("dias")
                        end if
                    tt.close: set tt = nothing
                c.close: set c = nothing
            end function

            function NombrePlantilla(presupuesto)
                dim c, sqlString, tt 

                sqlString = "select Nombre " & _
                            "from pre_Presupuesto_Encabezado " & _
                            "where presupuesto = '" & presupuesto & "' " & _
                            "and Usuario = '" & request.Cookies("Usuario") & "';"

                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set tt = c.execute(sqlString)
                        if (tt.bof or tt.eof) then
                            NombrePlantilla = "Nuevo Presupuesto"
                        else
                            NombrePlantilla = tt("Nombre")
                        end if
                    tt.close: set tt = nothing
                c.close: set c = nothing
            end function 

            function NombreMes(mes)     
                dim c, sqlString, tt 

                sqlString = "SELECT Nombre " & _
                              "FROM seg_cripto_Secuencias " & _
                             "WHERE (Tipo = 'M') " & _
                               "AND (CAST(Valor AS numeric(2)) = " & mes & ");"

                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set tt = c.execute(sqlString)
                        if (tt.bof or tt.eof) then
                            NombreMes = ""
                        else
                            NombreMes = tt("Nombre")
                        end if
                    tt.close: set tt = nothing
                c.close: set c = nothing
            end function

            Function FechaServer(fechaFormulario)
                Dim partes
                partes = Split(fechaFormulario, "/")
                
                FechaServer = partes(2) & "-" & _
                              Pad2(partes(1)) & "-" & _
                              Pad2(partes(0))
            End Function

            function fHasta(fechaDesde, Modelo)  
                dim c, sqlString, tt, res, d, m, a

                sqlString = "SELECT dbo.pre_Enc_Modelo_Hasta('" & request.Cookies("Usuario") & "', '" & Modelo & "', '" & FechaServer(fechaDesde) & "') AS Res;"

                set c = Server.CreateObject("ADODB.Connection")
                    c.open Application("Conn")
                    set tt = c.execute(sqlString)
                        if (tt.bof or tt.eof) then
                            fHasta = NULL
                        else
                            res = tt("Res")
                            partes = Split(res, "-")
                            fHasta = partes(2) & "/" & partes(1) & "/" & partes(0)
                        end if
                    tt.close: set tt = nothing
                c.close: set c = nothing  
            end function

            function recalcularHasta(Desde, Modelo, Mes, Amo)
                dim c, sqlString, tt, res, d, m, a, fDesde

                fdesde = Amo & "-" & right("00" & Mes, 2) & "-" & right("00" & right(Desde, 2), 2)
                sqlString = "SELECT dbo.pre_Enc_Modelo_Hasta('" & request.Cookies("Usuario") & "', '" & Modelo & "', '" & fdesde & "') AS Hasta;"

                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set tt = c.execute(sqlString)
                        if (tt.bof or tt.eof) then
                            recalcularHasta = NULL
                        else
                            res = tt("Hasta")
                            partes = Split(res, "-")

                            recalcularHasta = partes(2) & "/" & partes(1) & "/" & partes(0)
                        end if
                    tt.close: set tt = nothing
                c.close: set c = nothing  
            end function
        %>            
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <%
            dim con, t, sqlString, localUsuario
            dim cMes, cAmo, cPlantilla, cDesde, cHasta
            dim cOrigen, cDestino, cNombre, cAnterior
            dim cReglas, incActual, tpActual, ff, dIncremento
            
            usu = Request.Cookies("usuario")
            localUsuario = MonedaUsuario()

            '
            ' Abrimos la conexion con los datos...
            '

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            '
            ' Valores re-entrantes
            '
            cMes = Request.QueryString("m")

            if cMes = "" then
                cMes = Month(Date())
                cAmo = Year(Date())
                cPlantilla = "*"
                cDesde = FechaHoy()
                cHasta = FechaFinal()
                cNombre = "Nuevo Presupuesto"
                cOrigen = MonedaUsuario()
                cDestino = MonedaUsuario()
                cAnterior = Request.QueryString("p")
                cReglas = 0
            else
                cAmo = Request.QueryString("a")
                cPlantilla = Request.QueryString("p")
                cDesde = Request.QueryString("d")
                cHasta = Request.QueryString("h")
                cNombre = Request.QueryString("n")
                cOrigen = Request.QueryString("o1")
                cDestino = Request.QueryString("o2")
                cAnterior = Request.QueryString("pa")
                cReglas = Request.QueryString("r")

                if cPlantilla <> "*" then
                    '
                    ' Verificamos los campos y creamos el nuevo nombre
                    ' si es necesario...
                    '
                    set cbox = con.execute("SELECT * FROM pre_Presupuesto_Encabezado " & _
                                            "WHERE Usuario = '" & usu & "' " & _
                                            "AND Tipo = 'M' " & _
                                            "AND Presupuesto = '" & cPlantilla & "';")

                    if not (cbox.bof or cbox.eof) then
                        cHasta = fHasta(cDesde, cPlantilla)
                        cNombre = cAmo & "-" & RIGHT("0" & cMes, 2) & " " & NombrePlantilla(cPlantilla)
                        cNombre = cNombre & " de " & NombreMes(cMes)

                        cReglas = 1
                    end if
                    
                    cbox.close: set cbox = nothing
                end if
            end if
        %>  

        <div class="page-title-bar">
            <%= PageTitle %>
        </div>

        <form name="formulario" id="formulario" method="post" action="pre_enc_nuevo_crear.asp">
            <main>
                <br />

                <div class="contenedor">
                    <div class="line">
                        <label class="label normal">Fecha</label>
                        <select class="field normal" name="mes" id="mes" onChange="CambiarMes()">
                            <option value="1"  <% if cMes = 1  then response.write "selected='selected'" %>>Enero</option>
                            <option value="2"  <% if cMes = 2  then response.write "selected='selected'" %>>Febrero</option>
                            <option value="3"  <% if cMes = 3  then response.write "selected='selected'" %>>Marzo</option>
                            <option value="4"  <% if cMes = 4  then response.write "selected='selected'" %>>Abril</option>
                            <option value="5"  <% if cMes = 5  then response.write "selected='selected'" %>>Mayo</option>
                            <option value="6"  <% if cMes = 6  then response.write "selected='selected'" %>>Junio</option>
                            <option value="7"  <% if cMes = 7  then response.write "selected='selected'" %>>Julio</option>
                            <option value="8"  <% if cMes = 8  then response.write "selected='selected'" %>>Agosto</option>
                            <option value="9"  <% if cMes = 9  then response.write "selected='selected'" %>>Septiembre</option>
                            <option value="10" <% if cMes = 10 then response.write "selected='selected'" %>>Octubre</option>
                            <option value="11" <% if cMes = 11 then response.write "selected='selected'" %>>Noviembre</option>
                            <option value="12" <% if cMes = 12 then response.write "selected='selected'" %>>Diciembre</option>
                        </select>   
                    </div>

                    <div class="line">
                        <label class="label normal">Año</label>
                        <input class="field tiny" id="amo" name="amo" type="number" value="<%= cAmo %>" placeholder="0000"  />
                    </div>

                    <div class="line">
                        <label class="label normal">Plantilla</label>
                        <%
                            sqlString = "SELECT q.CodigoPlantilla, q.NombrePlantilla,	q.TipoPlantilla, q.IncrementoTotal, q.Desde, q.Hasta " & _
                                        "FROM (" & _
                                                "SELECT '*' AS CodigoPlantilla,'Presupuesto Vacio' AS NombrePlantilla," & _
                                                        " 0 AS Ordenamiento,	0 AS TipoPlantilla,	0 AS IncrementoTotal, NULL AS Desde, NULL AS Hasta " & _
                                                "FROM pre_Presupuesto_Encabezado AS d " & _
                                                "WHERE (d.Usuario = '" & usu & "') AND (d.Estatus = 1) AND Tipo = 'P'" & _                                      
                                                "UNION " & _
                                                "SELECT m.Presupuesto AS CodigoPlantilla, m.Nombre AS NombrePlantilla," & _
                                                        " 2 AS Ordenamiento,	1 AS TipoPlantilla, " & _
                                                        " dbo.pre_DiasModelo(m.Usuario, m.Presupuesto) AS IncrementoTotal, " & _
                                                        " m.Desde, m.Hasta " & _
                                                "FROM pre_Presupuesto_Encabezado AS m " & _
                                                "WHERE (m.Usuario = '" & usu & "') AND (m.Estatus = 1) AND Tipo = 'M'" & _
                                                ") AS q " & _
                                        "ORDER BY q.Ordenamiento,	q.NombrePlantilla;"

                            set cbox = con.execute(sqlString)

                                if (cbox.bof or cbox.eof) then
                                    response.write "<select class='field xxl' name='Plantilla' id='Plantilla' onChange='recargar()'>"
                                        response.write "<option value='*' selected='selected'>Presupuesto Vacio</option>"            
                                    response.write "</select>"

                                    response.write "<input id='incremento'  name='incremento' value='0'                   class='no-ver'/>"
                                    response.write "<input id='tipo'        name='tipo'       value='0'                   class='no-ver'/>"
                                    response.write "<input id='nPlantilla'  name='nPlantilla' value='Presupuesto Vacio'   class='no-ver'/>"
                                    response.write "<input id='npDesde'     name='npDesde' value='NULL'                   class='no-ver'/>"
                                    response.write "<input id='npHasta'     name='npHasta' value='NULL'                   class='no-ver'/>"
                                else
                                    response.write "<select class='field xxl' name='Plantilla' id='Plantilla' onChange='recargar()'>"

                                    Do
                                        response.write "<option value='" & cbox("CodigoPlantilla") & "'"

                                        if cbox("CodigoPlantilla") = cPlantilla then
                                            response.write " selected='selected'"

                                            tpActual = cbox("TipoPlantilla")
                                            incActual = cbox("IncrementoTotal")
                                            nomPlantilla = cbox("NombrePlantilla")

                                            if not isnull(cbox("Desde")) then cDesde = right(cbox("Desde"), 2) & "/" & RIGHT("00" & cMes, 2) & "/" & cAmo

                                            if cPlantilla <> "*" then
                                                cHasta = recalcularHasta(cbox("Desde"), cPlantilla, cMes, cAmo)
                                            end if

                                            if incActual = "" then incActual = 0                                            
                                        end if

                                        response.write ">" & cbox("NombrePlantilla") & "</option>"

                                        cbox.MoveNext

                                    Loop Until cbox.eof

                                    response.write "</select>"

                                    response.write "<input id='incremento'  name='incremento' value='" & incActual & "'     class='no-ver'/>"
                                    response.write "<input id='tipo'        name='tipo'       value='" & tpActual & "'      class='no-ver'/>"
                                    response.write "<input id='nPlantilla'  name='nPlantilla' value='" & nomPlantilla & "'  class='no-ver'/>"
                                    response.write "<input id='npDesde'     name='npDesde'    value='" & cDesde & "'        class='no-ver'/>"
                                    response.write "<input id='npHasta'     name='npHasta'    value='" & cHasta & "'        class='no-ver'/>"
                                end if
                            cbox.close: set cbox = nothing
                        %>
                    </div>

                    <div class="line">
                        <label class="label normal">Desde</label>
                        <input class="field tiny" id="desde" name="desde" type="text" value="<%= cDesde %>" placeholder="dd/mm/aaaa" onChange="VerificarFechaDesde()" />
                    </div>

                    <div class="line">
                        <label class="label normal">Hasta</label>
                        <input class="field tiny" id="hasta" name="hasta" type="text" value="<%= cHasta %>" placeholder="dd/mm/aaaa" onChange="VerificarFechaHasta()" />
                    </div>

                    <div class="line">
                        <label class="label normal">Nuevo</label>
                        <input class="field xxl" id="nuevoPre" name="nuevoPre" type="text" value="<%= cNombre %>" />
                    </div>

                    <div class="line">
                        <label class="label normal">Local Origen</label>                
                        <select class="field large" name="monedaOrigen" id="monedaOrigen">
                            <%
                            sqlString = "SELECT [Local], Simbolo + '  (' + NombreListas + ')' AS NombreLocal " & _
                                            "FROM seg_Cripto_NumParse_Locales " & _
                                        "WHERE [Local] <> 'NUM' " & _
                                        "ORDER BY NombreLocal ASC;"

                            set cbox = con.execute(sqlString)

                            if not (cbox.bof or cbox.eof) then
                                Do
                                    response.write "<option value='" & cbox("Local") & "' "
                                        if cOrigen = cbox("Local") then 
                                            response.write " selected='selected'" 
                                        end if
                                    response.write ">" & cbox("NombreLocal") & "</option>"

                                    cbox.MoveNext
                                Loop Until cbox.eof
                            end if

                            cbox.close: set cbox = nothing
                            %>
                        </select>
                    </div>

                    <div class="line">
                        <label class="label normal">Local Destino</label>                    
                        <select class="field large" name="monedaDestino" id="monedaDestino">
                            <%
                                sqlString = "SELECT [Local], Simbolo + '  (' + NombreListas + ')' AS NombreLocal " & _
                                                "FROM seg_Cripto_NumParse_Locales " & _
                                            "WHERE [Local] <> 'NUM' " & _
                                            "ORDER BY NombreLocal ASC;"

                                set cbox = con.execute(sqlString)

                                if not (cbox.bof or cbox.eof) then
                                    Do
                                        response.write "<option value='" & cbox("Local") & "' "
                                            if cDestino = cbox("Local") then 
                                                response.write " selected='selected'" 
                                            end if
                                        response.write ">" & cbox("NombreLocal") & "</option>"

                                        cbox.MoveNext
                                    Loop Until cbox.eof
                                end if

                                cbox.close: set cbox = nothing
                            %>
                        </select>                
                    </div>

                    <div class="line">
                        <label class="label normal">Cerrar</label>
                        <input class="no-ver" id="preAnterior" name="preAnterior" type="text" value="<%= cAnterior %>"/>
                        <select class="field xxl" name="preAnterior2" id="preAnterior2" disabled="true">
                            <%
                                sqlString = "SELECT Presupuesto, Nombre, Estatus " & _
                                            "FROM (" & _
                                                    "SELECT '*' AS Presupuesto, '' AS Nombre, 1 AS Estatus, 0 AS Ordenamiento " & _
                                                    "FROM pre_Presupuesto_Encabezado " & _
                                                    "UNION " & _
                                                    "SELECT Presupuesto, Nombre, Estatus, 1 AS Ordenamiento " & _
                                                    "FROM pre_Presupuesto_Encabezado " & _
                                                    "WHERE (Estatus = 1) " & _
                                                    "AND (Tipo <> 'M') " & _
                                                    "AND (Usuario = '" & usu & "') " & _
                                                ") AS q " & _
                                            "ORDER BY Ordenamiento ASC, Nombre ASC; "

                                set cbox = con.execute(sqlString)

                                if not (cbox.bof or cbox.eof) then
                                    Do
                                        response.write "<option value='" & cbox("Presupuesto") & "'"
                                            if cAnterior <> "" then
                                                if cAnterior = cbox("Presupuesto") then 
                                                    response.write " selected='selected'" 
                                                end if                      
                                            end if
                                        response.write ">" & cbox("Nombre") & "</option>"

                                        cbox.MoveNext
                                    Loop Until cbox.eof
                                end if

                                cbox.close: set cbox = nothing
                            %>
                        </select>                
                    </div>

                    <div class="line">
                        <label class="label normal">Aplicar Reglas</label>
                        <select class="field large" name="reglas" id="reglas" >
                            <option value = "1" <% if cReglas = 1 then response.write " selected='selected'" %>>Aplicar Todas las Reglas</option>
                            <option value = "0" <% if cReglas = 0 then response.write " selected='selected'" %>>No Aplicar Ninguna Regla</option>
                        </select>
                    </td>
                <div>

                <br /><br />
            </main>
        </form>

        <footer class="footer-contextual">
            <button class="footer-button" type="button" aria-label="Home" onclick="irA('/forma/mobile')">
                <i class="fa-solid fa-house"></i>
            </button>
                    
            <button class="footer-button" type="button" aria-label="Volver" onclick="volver()">
                <i class="fas fa-undo-alt"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Volver" onclick="cerrar()">
                <i class="fas fa-lock"></i>
            </button>                 
        </footer>

        <script>
            function volver() {
                history.back();
            }

            function irA(vinculo) {
                window.location.href = vinculo;
            }            

            function cerrar() {
                Swal.fire({
                    title: "Confirmación",
                    html: "<strong>Realizar Cierre</strong>",
                    icon: "question",
                    showCancelButton: true,
                    confirmButtonText: "Cerrar",
                    cancelButtonText: "Cancelar",
                    confirmButtonColor: "#d97706",   // Naranja serio
                    cancelButtonColor: "#0d6efd",    // Azul tranquilo
                    reverseButtons: true
                }).then(function(result) {
                    if (result.isConfirmed) {
                       document.getElementById("formulario").requestSubmit();
                    }
                });                
            }

            function right(cadena, caracteres) {
                return cadena.slice(cadena.length - caracteres, cadena.length);
            }
            
            function left(cadena, caracteres) {
                return cadena.slice(0, caracteres - cadena.length);
            }

            function VerificarFechaDesde() {
                var p   = document.getElementById("Plantilla").value;
                var inc = parseInt(document.getElementById("incremento").value, 10);
                var d   = document.getElementById("desde").value;
                var h   = document.getElementById("hasta").value;

                function parseFecha(fechaStr) {
                    var dia = parseInt(fechaStr.substring(0, 2), 10);
                    var mes = parseInt(fechaStr.substring(3, 5), 10) - 1; // JS meses 0-11
                    var ano = parseInt(fechaStr.substring(6, 10), 10);
                    return new Date(ano, mes, dia);
                }

                function formatearFecha(fecha) {
                    var dia = ("0" + fecha.getDate()).slice(-2);
                    var mes = ("0" + (fecha.getMonth() + 1)).slice(-2);
                    var ano = fecha.getFullYear();
                    return dia + "/" + mes + "/" + ano;
                }

                var dFecha = parseFecha(d);

                if (p === "*") {
                    var hFecha = parseFecha(h);

                    if (dFecha > hFecha) {
                        var temp = dFecha;
                        dFecha = hFecha;
                        hFecha = temp;

                        document.getElementById("desde").value = formatearFecha(dFecha);
                        document.getElementById("hasta").value = formatearFecha(hFecha);
                    }
                } else {
                    dFecha.setDate(dFecha.getDate() + inc);
                    document.getElementById("hasta").value = formatearFecha(dFecha);
                }
                actualizarTitulo();
            }

            function VerificarFechaHasta() {

                var p   = document.getElementById("Plantilla").value;
                var inc = parseInt(document.getElementById("incremento").value, 10);
                var d   = document.getElementById("desde").value;
                var h   = document.getElementById("hasta").value;

                function parseFecha(fechaStr) {
                    var dia = parseInt(fechaStr.substring(0, 2), 10);
                    var mes = parseInt(fechaStr.substring(3, 5), 10) - 1; // meses 0-11
                    var ano = parseInt(fechaStr.substring(6, 10), 10);
                    return new Date(ano, mes, dia);
                }

                function formatearFecha(fecha) {
                    var dia = ("0" + fecha.getDate()).slice(-2);
                    var mes = ("0" + (fecha.getMonth() + 1)).slice(-2);
                    var ano = fecha.getFullYear();
                    return dia + "/" + mes + "/" + ano;
                }

                var dFecha = parseFecha(d);
                var hFecha = parseFecha(h);

                if (p === "*") {

                    if (dFecha > hFecha) {
                        var temp = dFecha;
                        dFecha = hFecha;
                        hFecha = temp;

                        document.getElementById("desde").value = formatearFecha(dFecha);
                        document.getElementById("hasta").value = formatearFecha(hFecha);
                    }

                } else {

                    hFecha.setDate(hFecha.getDate() - inc);
                    document.getElementById("desde").value = formatearFecha(hFecha);
                }

                actualizarTitulo();
            }

            function recargar() {
                var params = new URLSearchParams();

                params.append("m", document.getElementById("mes").value);
                params.append("a", document.getElementById("amo").value);
                params.append("p", document.getElementById("Plantilla").value);
                params.append("d", document.getElementById("desde").value);
                params.append("h", document.getElementById("hasta").value);
                params.append("n", document.getElementById("nuevoPre").value);
                params.append("o1", document.getElementById("monedaOrigen").value);
                params.append("o2", document.getElementById("monedaDestino").value);
                params.append("pa", document.getElementById("preAnterior").value);
                params.append("r", document.getElementById("reglas").value);

                window.location.href = "pre_det_cerrar3.asp?" + params.toString();
            }  

            function nombreMes(mes) {
                var meses = [
                    "", 
                    "Enero", "Febrero", "Marzo", "Abril",
                    "Mayo", "Junio", "Julio", "Agosto",
                    "Septiembre", "Octubre", "Noviembre", "Diciembre"
                ];

                var index = parseInt(mes, 10);

                return meses[index] || "";
            }

            function actualizarTitulo() {
                var titulo = document.getElementById("nPlantilla").value;
                var tipoPlantilla = document.getElementById("tipo").value;
                var mes = document.getElementById("mes").value;
                var amo = document.getElementById("amo").value;
                var desde = document.getElementById("desde").value;
                var res = "";
            
                if (tipoPlantilla == 1) {
                    mes = parseInt(desde.substring(3, 5));
                    amo = parseInt(right(desde, 4));

                    document.getElementById("mes").value = mes;
                    document.getElementById("amo").value = amo;
                };

                res += amo + "-" + right("0" + mes, 2) + " ";
                res += titulo + " de " + nombreMes(mes);

                document.getElementById("nuevoPre").value = res;
            }

            function CambiarMes() {
                var m = parseInt(document.getElementById("mes").value, 10);
                var d = document.getElementById("desde").value;

                function formatearFecha(fecha) {
                    var dia = ("0" + fecha.getDate()).slice(-2);
                    var mes = ("0" + (fecha.getMonth() + 1)).slice(-2);
                    var ano = fecha.getFullYear();
                    return dia + "/" + mes + "/" + ano;
                }

                var dia = parseInt(d.substring(0, 2), 10);
                var ano = parseInt(d.substring(6, 10), 10);

                // JS usa meses 0-11
                var nuevaFecha = new Date(ano, m - 1, dia);

                document.getElementById("desde").value = formatearFecha(nuevaFecha);

                VerificarFechaDesde();
            }

            mask(document.getElementById('desde'),  ['99/99/9999']);
            mask(document.getElementById('hasta'),  ['99/99/9999']);
            mask(document.getElementById('amo'),    ['9999']);            
        </script>

        <% con.close: set con = nothing %> 
        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>