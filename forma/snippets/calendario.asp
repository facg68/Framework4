<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    Snip_Width = 385
%>
<!-- #include virtual = "/core/includes/snippets.inc" -->


<!--
    Snippet de Ejemplo 01:
    Calendario Mensual Actual
-->

<style>
    .cal_Snip_center         { text-align: center; }
    .cal_Snip_tiene_datos    { color: rgb(0, 0, 0); font-weight: bold; }
    .cal_Snip_no_tiene_datos { color: rgba(135, 135, 135, 1); }

    .cal_Snip_espacio_vacio  { background-color: rgb(145, 145, 145); }
    .cal_Snip_dia_actual     { background-color: rgb(238, 190, 190); }
    .cal_Snip_dia_pasado     { background-color: rgb(217, 235, 255); }
    .cal_Snip_dia_futuro     { background-color: rgb(233, 250, 222); }    
    
    .cal_snip_main {
        max-width: 96%;
        margin: 0.5rem auto;
        padding: 0;
        background: transparent;
        border-radius: 0;
        box-shadow: none;
        display: flex;
        font-family: sans-serif;
    }  
</style>  

<%    
    function cal_Snip_PrimerDiaMes(Amo, Mes)
        dim f
        f = Amo & "-" & right("00" & Mes, 2) & "-01"
        cal_Snip_PrimerDiaMes = WeekDay(f)   
    end function

    function cal_Snip_Bisiesto(Amo)
        if (Amo Mod 4) = 0 then 
            cal_Snip_Bisiesto = True 
        else 
            cal_Snip_Bisiesto = False
        end if       
    end function

    function cal_Snip_DiasPorMes(Amo, Mes)
        select case Mes
            case 1, 3, 5, 7, 8, 10, 12
                cal_Snip_DiasPorMes = 31
            case 4, 6, 9, 11
                cal_Snip_DiasPorMes = 30
            case 2
                if cal_Snip_Bisiesto(Amo) then 
                    cal_Snip_DiasPorMes = 29
                else
                    cal_Snip_DiasPorMes = 28
                end if
        end select         
    end function        

    function cal_Snip_LineasPorMes(Inicio, diasPorMes)
        dim k, sem, resto

        sem = ((Inicio + diasPorMes - 1) / 7.0)
        resto = sem - round(sem)
        cal_Snip_LineasPorMes = round(sem) 

        if resto > 0 then
            cal_Snip_LineasPorMes = cal_Snip_LineasPorMes + 1
        end if
    end function    

    function cal_Snip_NombreMes(Mes)
        dim cc, tt, sqlString

        sqlString = "SELECT Nombre FROM seg_cripto_Secuencias WHERE (Tipo = 'M') AND (Valor = " & Mes & ");"

        set cc = Server.CreateObject("ADODB.Connection")
        cc.open Application("Conn")
            set tt = cc.execute(sqlString)
                cal_Snip_NombreMes = tt("Nombre")
            tt.close: set tt = nothing
        cc.close: set cc = nothing
    end function

    function cal_Snip_horaNumStr(HoraNum)
        dim cadena, hh, mm

        if HoraNum <> "" then
            cadena = right("00000" & cStr(HoraNum), 4)
            hh = left(cadena, 2)
            mm = right(cadena, 2)

            cal_Snip_horaNumStr = hh & ":" & mm
        else  
            cal_Snip_horaNumStr = "&nbsp;"
        end if
    end function

    sub cal_Snip_Calendario(AA, MM, Tipo)  
        dim cc, tt, sqlString, diaRef
        dim aaAmo, ssAmo, aaMes, ssMes
        dim mMes(42, 2), base

        '-- Preparar Variables --
            base = "/forma/desktop/apps/agenda/cal/cal_eventos.asp?"

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
            <table class="tabla tabla-carbon tabla-calendario">
                <thead>
                    <tr style="background-color: rgb(0, 0, 0); color: rgb(255, 255, 255);" style="border: none;">
                        <th class="cal_Snip_center" style="border: none;">Do</th>
                        <th class="cal_Snip_center" style="border: none;">Lu</th>
                        <th class="cal_Snip_center" style="border: none;">Ma</th>
                        <th class="cal_Snip_center" style="border: none;">Mi</th>
                        <th class="cal_Snip_center" style="border: none;">Ju</th>
                        <th class="cal_Snip_center" style="border: none;">Vi</th>
                        <th class="cal_Snip_center" style="border: none;">Sa</th>
                    </tr>
                </thead>

                <tbody>
                    <%
                        '-- Preparación de Acceso a Datos --
                            for k = 1 to 42
                                mMes(k, 1) = "*"
                                mMes(k, 2) = "*"
                            next   

                            Inicio = cal_Snip_PrimerDiaMes(AA, MM)
                            dxMes = cal_Snip_DiasPorMes(AA, MM)
                            lxMes = cal_Snip_LineasPorMes(inicio, dxMes)
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
                                        claseTD = ""
                                        
                                        if mMes(Indice, 1) = "*" then ' Dia vacio
                                            claseTD = claseTD & "cal_Snip_espacio_vacio"
                                        else
                                            if diaRef = diaEntero then ' Dia actual
                                                if mMes(Indice, 2) = 0 then
                                                    claseTD = claseTD &  "cal_Snip_dia_actual "
                                                else
                                                    claseTD = claseTD &  "cal_Snip_dia_actual cal_Snip_tiene_datos "
                                                end if
                                            else
                                                Select Case Tipo
                                                    Case "M"
                                                        if diaRef = diaEntero then ' Dia actual
                                                            claseTD =  claseTD &  "cal_Snip_dia_actual"
                                                        else
                                                            if diaEntero < diaRef then ' Dia viejo
                                                                claseTD = claseTD &  "cal_Snip_dia_pasado "
                                                            else  ' dia futuro
                                                                claseTD = claseTD &  "cal_Snip_dia_futuro " 
                                                            end if
                                                        end if

                                                        if mMes(Indice, 2) > 0 then
                                                            claseTD = claseTD & "cal_Snip_tiene_datos "
                                                        end if                                                      

                                                    Case "A"
                                                        if mMes(Indice, 2) = 0 then
                                                            claseTD = claseTD & "cal_Snip_dia_pasado "
                                                        else
                                                            claseTD = claseTD & "cal_Snip_dia_pasado cal_Snip_tiene_datos "
                                                        end if

                                                    Case Else
                                                        if mMes(Indice, 2) = 0 then
                                                            claseTD = claseTD & "cal_Snip_dia_futuro "                                
                                                        else
                                                            claseTD = claseTD & "cal_Snip_dia_futuro cal_Snip_tiene_datos "                                 
                                                        end if
                                                end Select
                                            end if
                                        end if
                                    '-- Fin Definir claseTD

                                    response.write "<td class='clickable " & claseTD & "' style='text-align: center; border: none;'>" 
                                        if mMes(Indice, 1) <> "*" then
                                            if mMes(Indice, 2) = 0 then
                                                vinculo = base & "d=" & mMes(Indice, 1) & "&m=" & MM & "&a=" & AA 
                                                %>
                                                    <div class="cal_Snip_no_tiene_datos" onclick="calendario_abrir('<%= vinculo %>')">
                                                        <%= mMes(Indice, 1) %>
                                                    </div>
                                                <%
                                            else
                                                vinculo = base & "d=" & mMes(Indice, 1) & "&m=" & MM & "&a=" & AA 
                                                %>
                                                    <div class="cal_Snip_tiene_datos" onclick="calendario_abrir('<%= vinculo %>')">
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
        <% 
    end sub   

    Amo = YEAR(NOW())
    Mes = MONTH(NOW())
%>    
<div class="cal_snip_main">
    <div class="tabla-wrapper">
        <% cal_Snip_Calendario amo, mes, "M" %>
    </div>
</div>  

<script>
    function calendario_init() {
        redimWindow("calendario", <%= Snip_Width %>)
    }

    function calendario_abrir(vinculo) {
        window.location.href = vinculo;
    }
</script>    