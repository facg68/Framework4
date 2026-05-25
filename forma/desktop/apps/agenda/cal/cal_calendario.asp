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

            .main { max-width: 94%; }

            /* Clases Para El Calendario */
                /* ===== Utilidades básicas ===== */
                .center { text-align: center; }
                .tiene_datos { color: #000; font-weight: bold; }
                .no_tiene_datos { color: rgba(120,120,120,1); }

                /* ===== Contenedor general ===== */
                .calendar-wrapper {
                    width: 100%;
                    overflow: hidden;
                }

                /* ===== Marco visual del calendario ===== */
                .cal-shell {
                    background: #e9edf2;
                    border-radius: 12px;
                    box-shadow: 0 4px 10px rgba(0,0,0,0.2);
                    overflow: hidden;
                    padding: 2px;
                }

                /* ===== Tabla mensual ===== */
                .cal-month {
                    width: 100%;
                    border-collapse: collapse;
                    background: transparent;
                    table-layout: fixed;   /* 👈 ESTA es la bala de plata */
                    margin: 0;
                    padding: 0;
                }

                /* Limpieza mínima de tabla */
                .cal-month th,
                .cal-month td {
                    padding: 2px;
                    margin: 0;
                }

                /* ===== Encabezados ===== */
                .cal-month thead th {
                    background: linear-gradient(#2f2f2f, #1c1c1c);
                    color: #fff;
                    font-weight: bold;
                    font-size: 12px;
                    padding: 6px 0;
                    border-right: 1px solid #444;
                }

                .cal-month thead th:last-child {
                    border-right: none;
                }

                /* ===== Celdas del calendario ===== */
                .cal-day {
                    vertical-align: top;
                    padding: 4px;
                    border-right: 1px solid rgba(0,0,0,0.05);
                    border-bottom: 1px solid rgba(0,0,0,0.05);
                    font-family: Tahoma, Arial, sans-serif;
                    font-size: 11px;
                    line-height: 13px;
                }

                /* Estados por significado */
                .cal-day.dia_pasado   { background-color: #e7f0fb; } /* azul suave */
                .cal-day.dia_futuro   { background-color: #e9f4ea; } /* verde suave */
                .cal-day.dia_actual   { background-color: #f7eaea; } /* rosado suave */
                .cal-day.espacio_vacio{ background-color: #dcdcdc; }

                /* Bordes externos */
                .cal-month tbody tr:last-child .cal-day {
                    border-bottom: none;
                }
                .cal-month tbody tr .cal-day:last-child {
                    border-right: none;
                }

                /* ===== Número del día ===== */
                .cal-dia-numero {
                    display: block;
                    font-size: 14px;
                    line-height: 14px;
                    text-align: right;
                    margin-bottom: 2px;
                }

                /* ===== CONTEXTO RESETEADO DEL DÍA (CLAVE) ===== */
                /* Aquí el calendario se comporta como F3 */
                .cal-dia-reset {
                    font-family: Tahoma, Arial, sans-serif;
                    font-size: 11px;
                    line-height: 13px;

                    margin: 0;
                    padding: 0;

                    border: none;
                    background: transparent;

                    letter-spacing: normal;
                    word-spacing: normal;
                    white-space: normal;
                }

                /* Inputs internos: comportamiento clásico */
                .cal-dia-reset input {
                    display: block;
                    width: 100%;
                    height: 13px;

                    font-family: inherit;
                    font-size: inherit;
                    line-height: inherit;

                    margin: 0;
                    padding: 0;

                    border: none;
                    background: transparent;

                    box-sizing: border-box;
                    white-space: nowrap;
                    overflow: hidden;
                    text-overflow: ellipsis;
                    cursor: pointer;
                }

                .cal-dia-pie {
                    display: block;
                    font-size: 11px;
                    line-height: 12px;
                    font-weight: bold;
                    text-align: left;
                    color: #a5a5a5ff;
                }    
            /* Fin: Clases Para El Calendario */        
        </style>

        <%
            dim cc, Usuario

            Usuario = Request.Cookies("Usuario")

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")

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
                dim tt, sqlString

                sqlString = "SELECT Nombre FROM seg_cripto_Secuencias WHERE (Tipo = 'M') AND (Valor = " & Mes & ");"

                set tt = cc.execute(sqlString)
                    NombreMes = tt("Nombre")
                tt.close: set tt = nothing
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

            function VinculoMes(Amo, Mes, Incremento)
                dim a, m

                a = cInt(amo)
                m = cInt(mes)

                if Incremento = "+" then
                    m = m + 1

                    if m > 12 then
                        m = 1
                        a = a + 1
                    end if
                else
                    m = m - 1

                    if m < 1 then
                        m = 12
                        a = a - 1
                    end if
                end if

                VinculoMes = "cal_calendario.asp?a=" & a & "&m=" & m
            end function

            sub Calendario(AA, MM, Tipo)  
                dim sqlString, diaRef
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
                    <div class="cal-shell">
                        <table class="cal-month">
                            <tbody>
                                <tr style="background-color: rgb(0, 0, 0); color: rgb(255, 255, 255);">
                                    <th class="center">Domingo</th>
                                    <th class="center">Lunes</th>
                                    <th class="center">Martes</th>
                                    <th class="center">Miércoles</th>
                                    <th class="center">Jueves</th>
                                    <th class="center">Viernes</th>
                                    <th class="center">Sábado</th>
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

                                        sqlString = "exec cal_Mes_Conteo '" & Usuario & "', " & AA & ", " & MM

                                        set tt = cc.execute(sqlString)                          
                                            for k = 1 to dxMes
                                                mMes(Inicio, 1) = k
                                                mMes(Inicio, 2) = tt("Cuantos")
                                                
                                                Inicio = Inicio + 1
                                                tt.MoveNext
                                            next    
                                        tt.close: set tt = nothing

                                        Indice = 1
                                    '-- Fin: Preparación de Acceso a Datos --

                                    for j = 1 to lxMes
                                        response.write "<tr>"
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
                                                                            claseTD = claseTD & "dia_pasado "
                                                                        else  ' dia futuro
                                                                            claseTD = claseTD & "dia_futuro " 
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

                                                response.write "<td class='cal-day " & claseTD & "'>"
                                                    response.write "<div class='cal-dia-reset'>"                                               
                                                        if mMes(Indice, 1) <> "*" then
                                                            if mMes(Indice, 2) = 0 then
                                                                DibujarDia AA, MM, mMes(Indice, 1), 0
                                                            else
                                                                DibujarDia AA, MM, mMes(Indice, 1), 1
                                                            end if
                                                        end if
                                                    response.write "</div>"
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

            sub DibujarDia(Amo, Mes, Dia, Datos)
                dim f, cont, txt_mensaje, txt_Resto, ddSqlString, limite, k

                response.write "<table class='tabla-transparente cal-events'>"
                    response.write "<tr>"
                        response.write "<td class='cal-event'>"

                            mClass = "no_tiene_datos"
                            if Datos = 1 then mClass = "tiene_datos"                           
                            vinculo = "cal_eventos.asp?d=" & Dia & "&m=" & Mes & "&a=" & Amo

                            %>
                                <div class="cal-dia-numero <%= mClass %>" onclick="abrir('<%= vinculo %>')">
                                    <%= Dia %>
                                </div>                                
                            <%
                        response.write "</td>"
                    response.write "</tr>"                

                    f = amo & "-" & right("00" & mes, 2) & "-" & right("00" & dia, 2)

                    cont = 0
                    txt_Resto = ""
                    txt_mensaje = ""
                    ddSqlString = "exec dbo.Cal_Dia_Detalle '" & Usuario & "', '" & f & "'"

                    limite = ParametroUsuario("agenda", "cal_lineas_dia")
                    
                    if Isnull(limite) = True then
                        limite = 4
                    else
                        limite = cInt(limite)
                    end if

                    set tdia = cc.execute(ddSqlString)
                        if not (tdia.bof or tdia.eof) then      
                            Do
                                cont = cont + 1
                                txt_mensaje = horaNumStr(tdia("hora")) & "&nbsp;" & tdia("Descripcion")

                                if cont <= limite then
                                    response.write "<tr>"
                                        response.write "<td class='cal-event'>"
                                            %>
                                                <input style="width: 100%; color: rgb(0, 0, 0); background-color: transparent;"
                                                        type="text" readonly
                                                        value="<%= "&nbsp;" & txt_mensaje %>" 
                                                        title="<%= txt_mensaje %>"
                                                        onclick="AbrirObjeto('<%= tdia("tipo") %>', '<%= tdia("llave") %>')">
                                            <%
                                        response.write "</td>"                  
                                    response.write "</tr>"
                                else  
                                    txt_Resto = txt_Resto & txt_mensaje & "&#10;"
                                end if

                                tdia.MoveNext
                            Loop Until tdia.eof

                            if (cont < limite) then
                                for k = 1 to (limite - cont)
                                    response.write "<tr><td>&nbsp;</td></tr>"      
                                next
                            end if

                            if (cont = int(limite)) then
                                response.write "<tr><td>&nbsp;</td></tr>"
                            end if

                            if cont > limite then 
                                %>
                                    <tr>
                                        <td title="<%= txt_Resto %>">
                                            <div class="cal-dia-pie">
                                                <%= (cont - limite) %>+
                                            </div>                                        
                                        </td>
                                    </tr>
                                <%
                            end if
                        else
                            for kl = 1 to limite + 1
                                response.write "<tr><td>&nbsp;</td></tr>"
                            next 
                        end if
                    tdia.close: set tdia = nothing
                response.write "</table>"
            end sub    

            function FechaFormulario(Amo, Mes) 
                FechaFormulario = "01/" & RIGHT("0" & Mes, 2) & "/" & Amo
            end function

            Sub Filtro(Amo, Mes)
                dim tt, sqlString

                sqlString = "SELECT Nombre, Valor FROM seg_cripto_Secuencias WHERE (Tipo = 'M') ORDER BY CAST(Valor AS Numeric(2));"
            
                %>
                    <select class="campo" syle="width: 215px;" name="cboMes" id="cboMes" onChange="Requery()">
                        <%
                            set tt = cc.execute(sqlString)
                                if not (tt.bof or tt.eof) then
                                    Do
                                        response.write "<option value='" & tt("Valor") & "' "
                                            if cInt(tt("Valor")) = cInt(Mes) then 
                                                response.write " selected" 
                                            end if
                                        response.write ">" & tt("Nombre") & "</option>"

                                        tt.MoveNext
                                    Loop Until tt.eof
                                end if
                            tt.close: set tt = nothing                                
                        %>                        
                    </select>

                    <input class="campo" 
                           style="width: 60px; text-align: center;" 
                           id="txtAmo" name="txtAmo" type="text" value="<%= Amo %>" 
                           placeholder="aaaa" required 
                           onChange="Requery()"
                    />
                <%
            end Sub
        %>                       
    </head>

    <body plantilla="normal" reserva="125">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->
        <%
            dim amo, lineas, columnas, mes, tipo, am_Ref, am_Draw

            amo = request.QueryString("a")
            mes = request.QueryString("m")

            if amo = "" then amo = Year(Now())
            if mes = "" then mes = Month(Now())

            am_Ref = cdbl(Year(now()) & Right("00" & Month(now()), 2))        
        %>

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 30%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                <%= NombreMes(mes) & ", " & amo %>
            </div>
            
            <div style="flex: 0 0 70%; text-align: right;">
                Ir A:&nbsp;<% Filtro amo, mes %>

                &nbsp;&nbsp;

                <button class='form-btn tiny violeta' 
                        type='button' 
                        onclick="abrir('<%= VinculoMes(amo, mes, "-") %>')" >
                    <<
                </button>

                <button class='form-btn small violeta' 
                        type='button' 
                        onclick="abrir('cal_calendario.asp?a=<%= Year(Now()) %>&m=<%= Month(Now()) %>')" >
                    Hoy
                </button>                

                <button class='form-btn tiny violeta' 
                        type='button' 
                        onclick="abrir('<%= VinculoMes(amo, mes, "+") %>')" >
                    >>
                </button>

                &nbsp;&nbsp;

                <button class='form-btn normal azul' 
                        type='button' 
                        onclick="abrir('cal_semanal.asp?f=<%= FechaFormulario(Amo, Mes) %>')" >
                    Semanal
                </button>

                <button class='form-btn normal azul' 
                        type='button' 
                        onclick="abrir('cal_anual.asp?a=<%= Amo %>')" >
                    Anual
                </button>                                    
            </div>
        </div>        

        <div class="main main-scroll">
            <%
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
            %>
        </div>

        <script type="text/javascript">
            function abrir(vinculo) {
                window.location.href = vinculo;
            }

            function Requery() {
                verificarFiltroAmo();

                var mes = document.getElementById("cboMes").value;
                var amo = document.getElementById("txtAmo").value;
                var vinculo = "cal_calendario.asp?a=" + amo + "&m=" + mes;

                window.location.href = vinculo;
            }

            function verificarFiltroAmo() {
                var campo = document.getElementById("txtAmo");

                if (!campo) return; 

                if (campo.value.trim() === "") {
                    campo.value = <%= Amo %>;
                }
            }            

            function AbrirObjeto(tipo, llave) {
                if (tipo == "pre") {
                    var vinculo = "../pre/presupuestos/pre_det_tra_editar.asp?llaveCal=" + llave;
                    window.location.href = vinculo;
                };

                if (tipo == "con") {
                    var vinculo = "../cont/cont_editar.asp?con=" + llave;
                    window.location.href = vinculo;
                };

                if (tipo == "cal") {
                    var vinculo = "cal_eventos_editar.asp?o=m&s=" + llave;
                    window.location.href = vinculo;
                };        
            }       
            
            mask(document.getElementById('txtAmo'), ['9999']);
        </script>

        <% cc.close: set cc = nothing %>
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>