<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Calendario Anual" %>
        <title><%= PageTitle %></title>

        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0020"
            SysLockOut


            ' Funciones y Procedimientos ---------------------------------------------------------------
                function PrimerDiaMes(Amo, Mes)
                    dim f
                    f = Amo & "-" & right("00" & Mes, 2) & "-01"
                    PrimerDiaMes = WeekDay(f)   
                end function

                function Bisiesto(Amo)
                    if (Amo Mod 4) = 0 then 
                        Bisiesto = True 
                    else 
                        Bisiesto = False
                    end if       
                end function

                function DiasPorMes(Amo, Mes)
                    select case Mes
                        case 1, 3, 5, 7, 8, 10, 12
                            DiasPorMes = 31
                        case 4, 6, 9, 11
                            DiasPorMes = 30
                        case 2
                            if Bisiesto(Amo) then 
                                DiasPorMes = 29
                            else
                                DiasPorMes = 28
                            end if
                    end select         
                end function        

                function LineasPorMes(Inicio, diasPorMes)
                    dim k, sem, resto

                    sem = ((Inicio + diasPorMes - 1) / 7.0)
                    resto = sem - round(sem)
                    LineasPorMes = round(sem) 

                    if resto > 0 then
                        LineasPorMes = LineasPorMes + 1
                    end if
                end function    

                function NombreMes(Mes)
                    dim cc, tt, sqlString

                    sqlString = "SELECT Nombre FROM seg_cripto_Secuencias WHERE (Tipo = 'M') AND (Valor = " & Mes & ");"

                    set cc = Server.CreateObject("ADODB.Connection")
                    cc.open Application("Conn")
                        set tt = cc.execute(sqlString)
                            NombreMes = tt("Nombre")
                        tt.close: set tt = nothing
                    cc.close: set cc = nothing
                end function

                function horaNumStr(HoraNum)
                    dim cadena, hh, mm

                    if HoraNum <> "" then
                        cadena = right("00000" & cStr(HoraNum), 4)
                        hh = left(cadena, 2)
                        mm = right(cadena, 2)

                        horaNumStr = hh & ":" & mm
                    else  
                        horaNumStr = "&nbsp;"
                    end if
                end function

                sub Calendario(AA, MM, Tipo)  
                    dim cc, tt, sqlString, diaRef
                    dim aaAmo, ssAmo, aaMes, ssMes
                    dim mMes(42, 2)

                    '-- Preparar Variables --
                        if mm = 1 then
                            aaMes = 12
                            aaAmo = aa -1
                        else
                            aaMes = mm -1
                            aaAmo = aa
                        end if

                        if mm = 12 then
                            ssMes = 1
                            ssAmo = aa + 1
                        else
                            ssMes = mm + 1
                            ssAmo = aa
                        end if
                    '-- Fin Preparar Variables --

                    %>
                        <div class="line">
                            <div class="calendario tabla-carbon">
                                <div class="cal-head">
                                    <%= NombreMes(MM) & " / " & AA %>
                                </div>

                                <div class="cal-dias-semana">
                                    <div>Do</div>
                                    <div>Lu</div>
                                    <div>Ma</div>
                                    <div>Mi</div>
                                    <div>Ju</div>
                                    <div>Vi</div>
                                    <div>Sa</div>
                                </div>
                            <div class="cal-grid">
                                <%
                                    ' Preparación de Datos ----------------------------------------------------------------------------
                                        for k = 1 to 42
                                            mMes(k, 1) = "*"
                                            mMes(k, 2) = "*"
                                        next   

                                        Inicio = PrimerDiaMes(AA, MM)
                                        dxMes  = DiasPorMes(AA, MM)
                                        lxMes  = LineasPorMes(inicio, dxMes)

                                        diaRef = cDbl(Year(Now()) & right("00" & Month(Now()), 2) & right("00" & Day(Now()), 2))

                                        sqlString = "exec cal_Mes_Conteo '" & Request.Cookies("Usuario") & "', " & AA & ", " & MM

                                        set cc = Server.CreateObject("ADODB.Connection")
                                        cc.open Application("Conn")

                                        set tt = cc.execute(sqlString)

                                        for k = 1 to dxMes
                                            mMes(Inicio, 1) = k
                                            mMes(Inicio, 2) = tt("Cuantos")

                                            Inicio = Inicio + 1
                                            tt.MoveNext
                                        next    

                                        tt.close: set tt = nothing
                                        cc.close: set cc = nothing 

                                        Indice = 1
                                    ' -------------------------------------------------------------------------------------------------

                                    ' Render de los 42 espacios -----------------------------------------------------------------------
                                        for j = 1 to lxMes
                                            for k = 1 to 7
                                                if mMes(Indice, 1) <> "*" then
                                                    diaEntero = cDbl(AA & right("00" & MM, 2) & right("00" & mMes(Indice, 1), 2))
                                                end if

                                                claseTD = "cal-dia center "

                                                if mMes(Indice, 1) = "*" then
                                                    claseTD = claseTD & "espacio_vacio"
                                                else
                                                    if diaRef = diaEntero then
                                                        if mMes(Indice, 2) = 0 then
                                                            claseTD = claseTD & "dia_actual "
                                                        else
                                                            claseTD = claseTD & "dia_actual tiene_datos "
                                                        end if
                                                    else
                                                        Select Case Tipo
                                                            Case "M"
                                                                if diaEntero < diaRef then
                                                                    claseTD = claseTD & "dia_pasado "
                                                                else
                                                                    claseTD = claseTD & "dia_futuro "
                                                                end if

                                                                if mMes(Indice, 2) > 0 then
                                                                    claseTD = claseTD & "tiene_datos "
                                                                end if

                                                            Case "A"
                                                                if mMes(Indice, 2) = 0 then
                                                                    claseTD = claseTD & "dia_pasado "
                                                                else
                                                                    claseTD = claseTD & "dia_pasado tiene_datos "
                                                                end if

                                                            Case Else
                                                                if mMes(Indice, 2) = 0 then
                                                                    claseTD = claseTD & "dia_futuro "
                                                                else
                                                                    claseTD = claseTD & "dia_futuro tiene_datos "
                                                                end if
                                                        End Select
                                                    end if
                                                end if

                                                response.write "<div class='" & claseTD & "'>"

                                                if mMes(Indice, 1) <> "*" then
                                                    vinculo = "cal_calendario.asp?d=" & mMes(Indice, 1) & "&m=" & MM & "&a=" & AA

                                                    if mMes(Indice, 2) = 0 then
                                        %>
                                                        <div class="no_tiene_datos"
                                                            onclick="irA('<%= vinculo %>')">
                                                            <%= mMes(Indice, 1) %>
                                                        </div>
                                        <%
                                                    else
                                        %>
                                                        <div class="tiene_datos"
                                                            onclick="irA('<%= vinculo %>')">
                                                            <%= mMes(Indice, 1) %>
                                                        </div>
                                        <%
                                                    end if
                                                end if

                                                response.write "</div>"

                                                Indice = Indice + 1

                                            next
                                        next
                                    ' -------------------------------------------------------------------------------------------------        
                                %>  

                                </div> <!-- cal-grid -->
                            </div>
                        </div>
                    <%
                end sub  

                function fechaNuevo(amo)
                    dim d, m

                    d = Right("00" & Day(Now()), 2)
                    m = Right("00" & Day(Now()), 2)

                    fechaNuevo = amo & "-" & m & "-" & d
                end function
            '-------------------------------------------------------------------------------------------
        %>    

        <style>
            .contenedor { width: 96%; }

            /* Clases para los Calendarios -----------------------------------------------------------*/

                .calendario {
                    width: 100%;
                    border-radius: 12px;
                    overflow: hidden;
                    box-shadow: 0 0 6px var(--tabla-borde);
                    background-color: var(--tabla-fondo-body);
                    font-family: system-ui, sans-serif;
                }

                .cal-head {
                    background-color: var(--tabla-fondo-head);
                    color: var(--tabla-texto-head);
                    text-align: center;
                    font-weight: 600;
                    padding: 0.75rem;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                }

                .cal-dias-semana {
                    display: grid;
                    grid-template-columns: repeat(7, 1fr);
                    background-color: #000;
                    color: #fff;
                    font-weight: 600;
                }

                .cal-dias-semana div {
                    text-align: center;
                    padding: 0.5rem 0;
                }

                .cal-grid {
                    display: grid;
                    grid-template-columns: repeat(7, 1fr);
                }

                .cal-dia {
                    min-height: 60px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    background-color: #ffffff;

                    font-family: Ruda;
                    font-size: 1.2rem;
                }

                .dia_actual { background-color: rgb(238, 190, 190); }
                .dia_pasado { background-color: rgb(217, 235, 255); }
                .dia_futuro { background-color: rgb(233, 250, 222); }
                .espacio_vacio { background-color: rgb(145, 145, 145); }
                .tiene_datos { color: #000; font-family: 'Ruda Bold'; font-size: 1.3rem; }
                .no_tiene_datos { color: rgb(145, 145, 145); }
                
            /*----------------------------------------------------------------------------------------*/
        </style>                     
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     
        <%
            dim amo, lineas, columnas, mes, tipo, am_Ref, am_Draw

            amo = request.QueryString("a")
            if amo = "" then amo = Year(Now())

            am_Ref = cdbl(Year(now()) & Right("00" & Month(now()), 2))        
        %>

        <div class="page-title-bar">
            <div class="title"><%= "Calendario " & amo %></div>
        </div>

        <main>
            <br />

            <div class="contenedor">
                <%
                    mes = 0

                    for mes = 1 to 12
                        response.write "<div class='cal-mes'>"
                            am_Draw = cdbl(Amo & Right("00" & mes, 2))

                            if (am_Ref = am_Draw) then
                                Tipo = "M"
                            else
                                if (am_Ref > am_Draw) then
                                    Tipo = "A"
                                else
                                    Tipo = "F"
                                end if
                            end if

                            Calendario amo, mes, tipo
                        response.write "</div>"
                    next
                %>
            <div>

            <br />
        </main>

        <footer class="footer-contextual">
            <button class="footer-button" type="button" aria-label="Home" onclick="irA('/forma/mobile')">
                <i class="fa-solid fa-house"></i>
            </button>
                    
            <button class="footer-button" type="button" aria-label="Volver" onclick="volver()">
                <i class="fa-solid fa-rotate-left"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Año Anterior" onclick="irA('cal_anual.asp?a=<%= (amo - 1) %>')">
                <i class="fa-solid fa-backward"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Nuevo Evento" onclick="nuevo()">
                <i class="fa-solid fa-plus"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Siguiente Año" onclick="irA('cal_anual.asp?a=<%= (amo + 1) %>')">
                <i class="fa-solid fa-forward"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Volver al Año Actual" onclick="irA('cal_anual.asp?a=<%= year(date()) %>')">
                <i class="fa-solid fa-bullseye"></i>
            </button>
        </footer>

        <script>
            function volver() {
                history.back();
            }

            function irA(vinculo) {
                window.location.href = vinculo;
            }

            function nuevo(fechaSQL) {
                var vinculo = "cal_eventos_editar.asp?f=<%= fechaNuevo(amo) %>&s=*";
                window.location.href = vinculo;
            }

            document.addEventListener("DOMContentLoaded", function () {
                const hoy = new Date();
                const mesNumero = hoy.getMonth(); // 0 = enero
                const meses = [
                    "ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
                    "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"
                ];

                const mesActualTexto = meses[mesNumero];
                const encabezados = document.querySelectorAll(".cal-head");

                encabezados.forEach(head => {
                    if (head.textContent.toUpperCase().includes(mesActualTexto)) {
                        head.closest(".calendario").scrollIntoView({
                            behavior: "smooth",
                            block: "start"
                        });
                    }
                });

            });      
        </script>

        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>