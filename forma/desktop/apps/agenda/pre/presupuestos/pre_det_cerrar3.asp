<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>

        <title>Crear Nuevo Presupuesto y Mover Saldos</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0090"
            SysLockOut
        %>       
    
        <style>
            td, th {
                padding: 5px;
                font-size: 16px;
            }    

            .ldetalle {
                background-color: rgb(230, 230, 230);
                color: black;        
            }

            .ldetalle_impar {
                background-color: rgb(240, 240, 240);
                color: black;        
            }      

            .top {
                background-color: rgb(71,71,71);
                color: white;
            }

            .foot {
                background-color: rgb(89,89,89);
                color: white;
            }

            .CeldaDetalle {
                border: 1px solid rgb(186, 216, 232);
                padding:5px;
            }

            .vbControl_Verde {
                background-color: rgb(224, 255, 204);
                padding: 5px;
                border: 1px solid rgb(199, 199, 199);
            }

            .vbControl_Azul {
                background-color: rgb(222, 236, 255);
                padding: 5px;
                border: 1px solid rgb(199, 199, 199);      
            }

            .vbControl_Rojo {
                background-color: rgb(255, 222, 222);
                padding: 5px;
                border: 1px solid rgb(199, 199, 199);
            }            

            .control-label {
                font-size: 14px;
            }

            .button-13 {
                background-color: #fff;
                border: 1px solid #d5d9d9;
                border-radius: 8px;
                box-shadow: rgba(213, 217, 217, .5) 0 2px 5px 0;
                box-sizing: border-box;
                color: #0f1111;
                cursor: pointer;
                display: inline-block;
                font-family: "Amazon Ember",sans-serif;
                font-size: 12px;
                line-height: 25px;
                padding: 0 3px 0 3px;
                position: relative;
                text-align: center;
                text-decoration: none;
                user-select: none;
                -webkit-user-select: none;
                -touch-action: manipulation;
                vertical-align: middle;   
            }

            .button-13:hover {
                background-color: #f7fafa;
            }

            .button-13:focus {
                border-color: #008296;
                box-shadow: rgba(213, 217, 217, .5) 0 2px 5px 0;
                outline: 0;
            }

            .borde {
                border: 1px solid rgb(156, 156, 156);        
            }    
        </style>      

        <%
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

            function limpiar(cadena)
                dim char, k, res

                res = ""

                if cadena <> "" then
                    for k = 1 to (len(trim(cadena)))
                        char = mid(cadena, k, 1)

                        select case asc(char)
                        case 225: char = "a"
                        case 193: char = "A"
                        case 233: char = "e"
                        case 232: char = "e"
                        case 201: char = "E"
                        case 237: char = "i"
                        case 205: char = "I"
                        case 243: char = "o"
                        case 211: char = "O"
                        case 250: char = "u"
                        case 218: char = "U"
                        case 209: char = "N"
                        case 241: char = "n"
                        end select

                        res = res  & char
                    next

                    limpiar = res
                end if
            end function   

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

        <div style="width: 95%; margin: auto;">
            <br />

            <form name="form_transaccion" id="form_transaccion" method="post" action="pre_enc_nuevo_crear.asp">
                <table style="width: 100%;">
                    <tr>
                        <td style="width: 45%; text-align: left; font-size: 18px;">
                            Crear Nuevo Presupuesto
                        </td>

                        <td style="width: 55%; text-align: right;">
                            <button class='form-btn verde xl' type='submit'>Crear Presupuesto y Cerrar</button>

                            <a href='../lista.asp'>
                                <button type='button' class='form-btn rojo normal'>Cancelar</button>
                            </a>
                        </td>                        
                    </tr>
                </table>

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

                                        response.write ">" & limpiar(cbox("NombrePlantilla")) & "</option>"

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
                                response.write ">" & limpiar(cbox("NombreLocal")) & "</option>"

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
                                response.write ">" & limpiar(cbox("NombreLocal")) & "</option>"

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

                                    response.write ">" & limpiar(cbox("Nombre")) & "</option>"
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
                </div>
            </form>
        </div>

        <br /><br />

        <script>
            function right(cadena, caracteres) {
                return cadena.slice(cadena.length - caracteres, cadena.length);
            }
            
            function left(cadena, caracteres) {
                return cadena.slice(0, caracteres - cadena.length);
            }

            function VerificarFechaDesde() {
                var p = document.getElementById("Plantilla").value;
                var inc = document.getElementById("incremento").value;
                var d =  document.getElementById("desde").value;
                var h =  document.getElementById("hasta").value;
                var dia, mes, amo, temp, ffinal;
                var dFecha, hFecha;        

                if ( p == "*" ) {
                    dia = d.substring(0, 2);
                    mes = d.substring(3, 5);
                    amo = d.substring(6, 10);

                    temp = amo + "-" + mes + "-" + dia;
                    dFecha = dayjs(temp);

                    dia = h.substring(0, 2);
                    mes = h.substring(3, 5);
                    amo = h.substring(6, 10);

                    temp = amo + "-" + mes + "-" + dia;
                    hFecha = dayjs(temp);

                    if (dFecha > hFecha) {
                        temp = dayjs(dFecha).format("DD/MM/YYYY");
                        dFecha = dayjs(hFecha).format("DD/MM/YYYY");
                        hFecha = temp;

                        document.getElementById("desde").value = dFecha;
                        document.getElementById("hasta").value = hFecha;
                    }
                }
                else {
                    dia = d.substring(0, 2);
                    mes = d.substring(3, 5);
                    amo = d.substring(6, 10);

                    //
                    // Convertir fecha en "fecha internacional"
                    //

                    temp = amo + "-" + mes + "-" + dia;
                    var ffinal = dayjs(temp).add(inc, "day").format("DD/MM/YYYY");
                    document.getElementById("hasta").value = ffinal;
                }

                actualizarTitulo();
            }

            function VerificarFechaHasta() {
                var p = document.getElementById("Plantilla").value;
                var inc = document.getElementById("incremento").value;
                var d =  document.getElementById("desde").value;
                var h =  document.getElementById("hasta").value;
                var dia, mes, amo, temp, ffinal;
                var dFecha, hFecha;

                if ( p == "*" ) {
                    dia = d.substring(0, 2);
                    mes = d.substring(3, 5);
                    amo = d.substring(6, 10);

                    temp = amo + "-" + mes + "-" + dia;
                    dFecha = dayjs(temp);

                    dia = h.substring(0, 2);
                    mes = h.substring(3, 5);
                    amo = h.substring(6, 10);

                    temp = amo + "-" + mes + "-" + dia;
                    hFecha = dayjs(temp);

                    if (dFecha > hFecha) {
                        temp = dayjs(dFecha).format("DD/MM/YYYY");
                        dFecha = dayjs(hFecha).format("DD/MM/YYYY");
                        hFecha = temp;

                        document.getElementById("desde").value = dFecha;
                        document.getElementById("hasta").value = hFecha;
                    }
                }
                else {
                    dia = h.substring(0, 2);
                    mes = h.substring(3, 5);
                    amo = h.substring(6, 10);

                    //
                    // Convertir fecha en "fecha internacional"
                    //

                    temp = amo + "-" + mes + "-" + dia;
                    var ffinal = dayjs(temp).subtract(inc, "day").format("DD/MM/YYYY");
                    document.getElementById("desde").value = ffinal;
                }
                actualizarTitulo();        
            }

            function recargar() {
                var m = document.getElementById("mes").value;
                var a = document.getElementById("amo").value; 
                var p = document.getElementById("Plantilla").value;
                var d = document.getElementById("desde").value; 
                var h = document.getElementById("hasta").value; 
                var n = document.getElementById("nuevoPre").value; 
                var o1 = document.getElementById("monedaOrigen").value; ;
                var o2 = document.getElementById("monedaDestino").value;
                var pa = document.getElementById("preAnterior").value; 
                var r = document.getElementById("reglas").value; 

                var v = "pre_det_cerrar3.asp";
                v += "?m=" + m;
                v += "&a=" + a; 
                v += "&p=" + p;
                v += "&d=" + d;
                v += "&h=" + h;
                v += "&n=" + n;
                v += "&o1=" + o1;
                v += "&o2=" + o2;
                v += "&pa=" + pa;
                v += "&r=" + r;

                // actualizarTitulo();

                window.location.href = v;
            }      

            function nombreMes(mes) {
                var res;

                switch(parseInt(mes)) {
                    case 1: res = "Enero"; break;
                    case 2: res = "Febrero"; break;
                    case 3: res = "Marzo"; break;
                    case 4: res = "Abril"; break;
                    case 5: res = "Mayo"; break;
                    case 6: res = "Junio"; break;
                    case 7: res = "Julio"; break;
                    case 8: res = "Agosto"; break;
                    case 9: res = "Septiembre"; break;
                    case 10: res = "Octubre"; break;
                    case 11: res = "Noviembre"; break;
                    case 12: res = "Diciembre"; break;
                };    

                return res;
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
                var dia, mes, amo, temp, ffinal;
                var m = document.getElementById("mes").value;
                var d = document.getElementById("desde").value;

                dia = d.substring(0, 2);
                mes = m;
                amo = d.substring(6, 10);

                temp = mes + "/" + dia + "/" + amo;
                temp = dayjs(temp).format("DD/MM/YYYY");
                document.getElementById("desde").value = temp;

                VerificarFechaDesde();            
            }

            mask(document.getElementById('desde'),  ['99/99/9999']);
            mask(document.getElementById('hasta'),  ['99/99/9999']);
            mask(document.getElementById('amo'),    ['9999']);
        </script>

        <% con.close: set con = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->            
    </body>
</html>