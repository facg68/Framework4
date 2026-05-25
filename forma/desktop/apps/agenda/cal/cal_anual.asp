<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->   
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0020"
            SysLockOut
        %>      

        <style>
            .main { max-width: 94%; }

            .center { text-align: center; }
            .tiene_datos { color: rgb(0, 0, 0); font-weight: bold; }
            .no_tiene_datos { color: rgba(135, 135, 135, 1); }

            .espacio_vacio  { background-color: rgb(145, 145, 145); }
            .dia_actual     { background-color: rgb(238, 190, 190); }
            .dia_pasado     { background-color: rgb(217, 235, 255); }
            .dia_futuro     { background-color: rgb(233, 250, 222); }

            .calendario-grid {
                display: grid;
                grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
                gap: 0.4rem;
                width: 100%;
            }            
           
            .cal-mes {
                padding: 0.3rem;
            }    

            .cal-mes table {
                width: 100%;
                max-width: 100%;
                table-layout: fixed;
                border-collapse: collapse;
            }   
            
            .cal-mes th,
            .cal-mes td {
                width: calc(100% / 7);
                box-sizing: border-box;
                text-align: center;
                padding: clamp(0.25rem, 0.4vw, 0.5rem);
                font-size: clamp(0.7rem, 1vw, 0.85rem);
                white-space: nowrap;
                border: none;
            }            
        </style>  

        <%
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
                '-- Fin: Preparar Variables --

                %>     
                    <div class="tabla-wrapper" style="width: 100%;">
                        <table class="tabla tabla-carbon tabla-calendario">
                            <thead>
                                <tr>
                                    <th colspan="7">
                                        <div onclick="abrir('<%= "cal_calendario.asp?a=" & AA & "&m=" & MM %>')">
                                            <%= NombreMes(MM) & " / " & AA %>
                                        </div>
                                    </th>
                                </tr>
                            </thead>

                            <tbody>
                                <tr style="background-color: rgb(0, 0, 0); color: rgb(255, 255, 255);">
                                    <th class="center">Do</th>
                                    <th class="center">Lu</th>
                                    <th class="center">Ma</th>
                                    <th class="center">Mi</th>
                                    <th class="center">Ju</th>
                                    <th class="center">Vi</th>
                                    <th class="center">Sa</th>
                                </tr>

                                <%
                                    '-- Preparación de Acceso a Datos --
                                        for k = 1 to 42
                                            mMes(k, 1) = "*"
                                            mMes(k, 2) = "*"
                                        next   

                                        Inicio = PrimerDiaMes(AA, MM)
                                        dxMes = DiasPorMes(AA, MM)
                                        lxMes = LineasPorMes(inicio, dxMes)
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
                                    '-- Fin: Preparación de Acceso a Datos --

                                    for j = 1 to lxMes
                                        response.write "<tr style='border: none;'>"
                                            for k = 1 to 7
                                                if mMes(Indice, 1) <> "*" then
                                                    diaEntero = cDbl(AA & right("00" & MM, 2) & right("00" & mMes(Indice, 1), 2))
                                                end if

                                                '-- Definir claseTD --
                                                    claseTD = "center "
                                                    
                                                    if mMes(Indice, 1) = "*" then ' Dia vacio
                                                        claseTD = claseTD & "espacio_vacio"
                                                    else
                                                        if diaRef = diaEntero then ' Dia actual
                                                            if mMes(Indice, 2) = 0 then
                                                                claseTD = claseTD &  "dia_actual "
                                                            else
                                                                claseTD = claseTD &  "dia_actual tiene_datos "
                                                            end if
                                                        else
                                                            Select Case Tipo
                                                                Case "M"
                                                                    if diaRef = diaEntero then ' Dia actual
                                                                        claseTD =  claseTD &  "dia_actual"
                                                                    else
                                                                        if diaEntero < diaRef then ' Dia viejo
                                                                            claseTD = claseTD &  "dia_pasado "
                                                                        else  ' dia futuro
                                                                            claseTD = claseTD &  "dia_futuro " 
                                                                        end if
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
                                                            end Select
                                                        end if
                                                    end if
                                                '-- Fin Definir claseTD

                                                response.write "<td class='" & claseTD & "' style='border: none;'>" 
                                                    if mMes(Indice, 1) <> "*" then
                                                        if mMes(Indice, 2) = 0 then
                                                            vinculo = "cal_eventos.asp?d=" & mMes(Indice, 1) & "&m=" & MM & "&a=" & AA 
                                                            %>
                                                                <div class="no_tiene_datos" onclick="abrir('<%= vinculo %>')">
                                                                    <%= mMes(Indice, 1) %>
                                                                </div>
                                                            <%
                                                        else
                                                            vinculo = "cal_eventos.asp?d=" & mMes(Indice, 1) & "&m=" & MM & "&a=" & AA 
                                                            %>
                                                                <div class="tiene_datos" onclick="abrir('<%= vinculo %>')">
                                                                    <%= mMes(Indice, 1) %>
                                                                </div>
                                                            <%
                                                        end if
                                                    end if
                                                response.write "</td>"

                                                Indice = Indice + 1   
                                            next  
                                        response.write "</tr>"
                                    next
                                %>                                                      
                            </tbody>
                        </table>
                    </div>
                <% 
            end sub            
        %>                       
    </head>

    <body plantilla="normal" reserva="125">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->
        <%
            dim amo, lineas, columnas, mes, tipo, am_Ref, am_Draw

            amo = request.QueryString("a")
            if amo = "" then amo = Year(Now())

            am_Ref = cdbl(Year(now()) & Right("00" & Month(now()), 2))        
        %>

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                Calendario <%= amo %>
            </div>
            
            <div style="flex: 0 0 50%; text-align: right;">
                <button class='form-btn normal violeta' 
                        type='button' 
                        onclick="abrir('<%= "cal_anual.asp?a=" & (amo - 1) %>')" >
                    < <%= (amo - 1) %>
                </button>

                <button class='form-btn normal violeta' 
                        type='button' 
                        onclick="abrir('<%= "cal_anual.asp?a=" & (amo + 1) %>')" >
                    <%= (amo + 1) %> >
                </button>

                &nbsp;

                <button class='form-btn normal verde' 
                        type='button' 
                        onclick="abrir('cal_calendario.asp')" >
                    Volver
                </button>                
            </div>
        </div>        

        <div class="main main-scroll">
            <div class="calendario-grid">
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
            </div>            
        </div>

        <script type="text/javascript">
            function abrir(vinculo) {
                window.location.href = vinculo;
            }
        </script>

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>