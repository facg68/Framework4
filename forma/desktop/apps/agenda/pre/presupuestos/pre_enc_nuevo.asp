<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>

        <title>Nuevo Presupuesto</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0090"
            SysLockOut

            '
            ' Funciones y Procedimientos
            '

            Function FechaHoy()
                dim a, m, d

                d = RIGHT("00" & day(Date) ,2)
                m = RIGHT("00" & month(Date), 2)
                A = year(Date)

                FechaHoy = d & "/" & m & "/" & a
            end function

            Function FechaFinal()
                dim a, m, d, f

                f = DateAdd("d", 15, Date())

                d = RIGHT("00" & day(f) ,2)
                m = RIGHT("00" & month(f), 2)
                A = year(f)

                FechaFinal = d & "/" & m & "/" & a
            end function      

            Function HoraHoy()
                dim h, m
                
                h = RIGHT("00" & Hour(Time()), 2)
                m = RIGHT("00" & Minute(Time()), 2)

                HoraHoy = h & ":" & m
            end function

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
                Select Case mes
                case 1: NombreMes = "Enero"
                case 2: NombreMes = "Febrero"
                case 3: NombreMes = "Marzo"
                case 4: NombreMes = "Abril"
                case 5: NombreMes = "Mayo"
                case 6: NombreMes = "Junio"
                case 7: NombreMes = "Julio"
                case 8: NombreMes = "Agosto"
                case 9: NombreMes = "Septiembre"
                case 10: NombreMes = "Octubre"
                case 11: NombreMes = "Noviembre"
                case 12: NombreMes = "Diciembre"
                End Select
            end function

            function FechaServer(fechaFormulario)
                dim d, m, a

                d = left(fechaFormulario, 2)
                m = mid(fechaFormulario, 4, 2)
                a = right(fechaFormulario, 4)

                FechaServer = a & "-" & right("0" & m, 2) & "-" & right("0" & d, 2)
            end function

            function FechaFormulario(fechaServer)
                dim p1, p2, p3, k, p, cadena, segmento, inicio

                '
                ' Parsear a mano... Uff!!!
                '

                p = 0
                cadena = FechaServer & "/"
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

                FechaFormulario = p2 & "/" & p1 & "/" & p3
            end function  

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

                d = right("00" & right(res, 2), 2)
                m = right("00" & mid(res, 6, 2), 2)
                a = left(res, 4)

                fHasta = d & "/" & m & "/" & a
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

                d = right("00" & right(res, 2), 2)
                m = right("00" & mid(res, 6, 2), 2)
                a = left(res, 4)

                recalcularHasta = d & "/" & m & "/" & a          
                end if

                tt.close: set tt = nothing
                c.close: set c = nothing  
            end function
        %>    
    </head>

      <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

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
                cAnterior = "*"
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

        <br />

        <form name="formulario" id="formulario" method="post" action="pre_enc_nuevo_crear.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Crear Nuevo Presupuesto
                </div>
                
                <div style="flex: 0 0 50%; text-align: right;">
                    <button type="button" class="form-btn verde large" onclick="submit()">Crear Presupuesto</button>    
                    <button type="button" class="form-btn rojo normal" onclick="volver()">Cancelar</button>    
                </div>
            </div>        

            <div class="main main-scroll">
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
                    <input class="field tiny" id="amo" name="amo" type="text" value="<%= cAmo %>" placeholder="0000" />
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
                                response.write "<select class='field large' name='Plantilla' id='Plantilla' onChange='recargar()'>"
                                    response.write "<option value='*' selected='selected'>Presupuesto Vacio</option>"            
                                response.write "</select>"

                                %>
                                    <div class="no-ver">
                                        <input id="incremento"  name="incremento"   value="0"                   />
                                        <input id="tipo"        name="tipo"         value="0"                   />
                                        <input id="nPlantilla"  name="nPlantilla"   value="Presupuesto Vacio"   />
                                        <input id="npDesde"     name="npDesde"      value="NULL"                />
                                        <input id="npHasta"     name="npHasta"      value="NULL"                />
                                    </div>
                                <%
                            else
                                response.write "<select class='field large' name='Plantilla' id='Plantilla' onChange='recargar()'>"
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

                               %>
                                    <div class="no-ver">
                                        <input id="incremento"  name="incremento"   value="<%= incActual %>" />
                                        <input id="tipo"        name="tipo"         value="<%= tpActual %>" />
                                        <input id="nPlantilla"  name="nPlantilla"   value="<%= nomPlantilla %>" />
                                        <input id="npDesde"     name="npDesde"      value="<%= cDesde %>" />
                                        <input id="npHasta"     name="npHasta"      value="<%= cHasta %>" />
                                    </div>
                                <%                                
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
                    <label class="label normal">Nombre</label>
                    <input class="field xl" id="nuevoPre" name="nuevoPre" type="text" value="<%= cNombre %>" />
                </div>

                <div class="line">
                    <label class="label normal">Moneda Origen</label>
                    <select class="field normal" name="monedaOrigen" id="monedaOrigen">
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
                    <label class="label normal">Moneda Destino</label>
                    <select class="field normal" name="monedaDestino" id="monedaDestino">
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
                    <label class="label normal">Auto-Cerrar</label>
                    <select class="field xxl" name="preAnterior" id="preAnterior">
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
                    <label class="label normal">Reglas</label>
                    <select class="field large" name="reglas" id="reglas" >
                        <option value = "1" <% if cReglas = 1 then response.write " selected='selected'" %>>Aplicar Todas las Reglas</option>
                        <option value = "0" <% if cReglas = 0 then response.write " selected='selected'" %>>No Aplicar Ninguna Regla</option>
                    </select>
                </div>                                                       
            </div>    
        </form>

        <script>
            function submit() {
                document.getElementById("formulario").submit();
            }

            function volver() {
                window.location.href = "../lista.asp";
            }

            function right(cadena, caracteres) {
                return cadena.slice(cadena.length - caracteres);
            }

            function left(cadena, caracteres) {
                return cadena.slice(0, caracteres);
            }

            function ajustarFechas(modo) {
                const plantilla  = getVal("Plantilla");
                const incremento = parseInt(getVal("incremento"), 10);
                const desdeStr   = getVal("desde");
                const hastaStr   = getVal("hasta");

                let fDesde, fHasta;

                if (plantilla === "*") {
                    fDesde = fecha_StrToDate(desdeStr);
                    fHasta = fecha_StrToDate(hastaStr);

                    if (fDesde > fHasta) {
                        [fDesde, fHasta] = [fHasta, fDesde];
                    }
                } else if (modo === "desde") {
                    fDesde = fecha_StrToDate(desdeStr);
                    fHasta = new Date(fDesde);
                    fHasta.setDate(fDesde.getDate() + incremento);
                } else if (modo === "hasta") {
                    fHasta = fecha_StrToDate(hastaStr);
                    fDesde = new Date(fHasta);
                    fDesde.setDate(fHasta.getDate() - incremento);
                }

                setVal("desde", fecha_Format(fDesde));
                setVal("hasta", fecha_Format(fHasta));

                actualizarTitulo();
            }

            const VerificarFechaDesde = () => ajustarFechas("desde");
            const VerificarFechaHasta = () => ajustarFechas("hasta");

            function recargar() {
                const params = {
                    m:  getVal("mes"),
                    a:  getVal("amo"),
                    p:  getVal("Plantilla"),
                    d:  getVal("desde"),
                    h:  getVal("hasta"),
                    n:  getVal("nuevoPre"),
                    o1: getVal("monedaOrigen"),
                    o2: getVal("monedaDestino"),
                    pa: getVal("preAnterior"),
                    r:  getVal("reglas")
                };

                const queryString = new URLSearchParams(params).toString();
                window.location.href = `pre_enc_nuevo.asp?${queryString}`;
            }

            function nombreMes(mes) {
                const meses = [
                    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
                ];
                return meses[parseInt(mes, 10) - 1] || null;
            }

            function actualizarTitulo() {
                let titulo        = getVal("nPlantilla");
                let tipoPlantilla = Number(getVal("tipo"));
                let mes           = Number(getVal("mes"));
                let amo           = Number(getVal("amo"));
                let desde         = getVal("desde");

                if (tipoPlantilla === 1 && desde.length >= 10) {
                    mes = Number(desde.substring(3, 5));
                    amo = Number(desde.slice(-4));
                    setVal("mes", mes);
                    setVal("amo", amo);
                }

                const mesFormateado = String(mes).padStart(2, "0");
                const res = `${amo}-${mesFormateado} ${titulo} de ${nombreMes(mes)}`;

                setVal("nuevoPre", res);
            }

            function CambiarMes() {
                const m = Number(getVal("mes"));
                const d = getVal("desde");
                const dia = d.substring(0, 2);
                const año = d.substring(6, 10);
                const fechaObj = new Date(Number(año), m - 1, Number(dia));

                const dd = String(fechaObj.getDate()).padStart(2, '0');
                const mm = String(fechaObj.getMonth() + 1).padStart(2, '0');
                const yyyy = fechaObj.getFullYear();

                const fechaFormateada = `${dd}/${mm}/${yyyy}`;
                setVal("desde", fechaFormateada);

                VerificarFechaDesde();
            }

            mask(document.getElementById('desde'), ['99/99/9999']);
            mask(document.getElementById('hasta'), ['99/99/9999']);
            mask(document.getElementById('amo'), ['9999']);
        </script>

        <% con.close: set con = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->       
    </body>
</html>