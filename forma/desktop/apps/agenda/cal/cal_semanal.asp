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
            
            .tabla {
                table-layout: fixed;
                width: 100%;
            }  
            
            td { 
                    vertical-align: top !important; 
                    font-family: 'Arial Narrow'; 
                    font-size: 14px;
                    text-align: left;
                    padding: 0px;
            }
        </style>  

        <%
            dim con, t, sqlString, PrimerDia, primeraVez

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

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

            Function Hoy()
                ' Llega en formato MM/dd/aaaa

                Dim d, m, a

                d = Day(Now())
                m = Month(Now())
                a = Year(Now())

                Hoy = Right("0" & d, 2) & "/" & _
                      Right("0" & m, 2) & "/" & _
                      a               
            End Function

            Function PrimerDiaSemanaFormulario(fechaTexto) 
                ' Llega en formato dd/MM/aaaa

                Dim partes, d, m, a, f, diaSemana

                partes = Split(fechaTexto, "/")
                d = CInt(partes(0))
                m = CInt(partes(1))
                a = CInt(partes(2))

                f = DateSerial(a, m, d)
                diaSemana = Weekday(f, vbSunday)
                f = DateAdd("d", -(diaSemana - 1), f)

                PrimerDiaSemanaFormulario = _
                    Right("0" & Day(f), 2) & "/" & _
                    Right("0" & Month(f), 2) & "/" & _
                    Year(f)
            End Function

            Function PrimerDiaSemanaSQL(fechaTexto) 
                ' Llega en formato dd/MM/aaaa

                Dim partes, d, m, a, f, diaSemana

                partes = Split(fechaTexto, "/")
                d = CInt(partes(0))
                m = CInt(partes(1))
                a = CInt(partes(2))

                f = DateSerial(a, m, d)
                diaSemana = Weekday(f, vbSunday)
                f = DateAdd("d", -(diaSemana - 1), f)

                PrimerDiaSemanaSQL = _
                    Year(f) & "-" & _
                    Right("0" & Month(f), 2) & "-" & _
                    Right("0" & Day(f), 2)
            End Function  

            Function Parse(texto)

                Dim partes, i, bloque
                Dim mensaje, meta, datos, llave, tipo
                Dim salida
                Dim pIni, pFin

                salida = ""

                ' Blindaje básico
                If IsNull(texto) Then
                    Parse = ""
                    Exit Function
                End If

                texto = CStr(texto)

                ' Normalizar separadores <br>
                texto = Replace(texto, "<br />", "<br/>", 1, -1, vbTextCompare)
                texto = Replace(texto, "<br>", "<br/>", 1, -1, vbTextCompare)

                ' Garantizar separador final
                If Trim(texto) <> "" Then
                    texto = texto & "<br/>"
                End If

                partes = Split(texto, "<br/>")

                For i = 0 To UBound(partes)
                
                    bloque = Trim(partes(i))
                    If bloque <> "" Then

                        pIni = InStr(bloque, "[[")
                        pFin = InStr(bloque, "]]")

                        If pIni > 0 And pFin > pIni Then

                            mensaje = Trim(Left(bloque, pIni - 1))
                            meta = Mid(bloque, pIni + 2, pFin - (pIni + 2))

                            datos = Split(meta, "|")
                            If UBound(datos) = 1 Then

                                llave = datos(0)
                                tipo  = datos(1)

                                salida = salida & _
                                    "<input class=""field frame"" style=""width:100%;"" " & _
                                    "type=""text"" readonly " & _
                                    "value=""" & mensaje & """ " & _
                                    "title=""" & mensaje & """ " & _
                                    "onclick=""abrir('" & tipo & "','" & llave & "')"">" & _
                                    "<br/>" & vbCrLf
                            End If
                        End If
                    End If
                Next

                Parse = salida

            End Function

            Function NuevaFecha(fechaTexto, dias)
                Dim p, f

                ' Separar dd/MM/yyyy
                p = Split(fechaTexto, "/")

                ' Crear fecha SIN depender del formato regional
                f = DateSerial(CInt(p(2)), CInt(p(1)), CInt(p(0)))

                ' Sumar o restar días
                f = DateAdd("d", dias, f)

                ' Devolver siempre dd/MM/yyyy
                NuevaFecha = Right("0" & Day(f), 2) & "/" & _
                             Right("0" & Month(f), 2) & "/" & _
                             Year(f)
            End Function    

            Function NuevoDia(fechaTexto, dias)
                Dim p, f, d

                p = Split(fechaTexto, "/")

                f = DateSerial(CInt(p(2)), CInt(p(1)), CInt(p(0)))
                f = DateAdd("d", dias, f)

                d = Day(f)

                NuevoDia = d
            End Function      

            Function VinculoEventos(FechaFormulario, dias)
                dim p, f

                p = Split(FechaFormulario, "/")
                f = DateSerial(CInt(p(2)), CInt(p(1)), CInt(p(0)))
                f = DateAdd("d", dias, f)                

                VinculoEventos = "cal_eventos.asp?o=s&d=" & Day(f) & "&m=" & Month(f) & "&a=" & Year(f)
            End Function
        %>                       
    </head>

    <body plantilla="tabla" reserva="125" onload="iniciar()">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->
        <%
            '-- Preparar Datos

                diaFormulario = request.QueryString("d") 'Viene en formato "dd/MM/aaaa"

                if diaFormulario = "" then 
                    diaFormulario = PrimerDiaSemanaFormulario(Hoy())
                else
                    diaFormulario = PrimerDiaSemanaFormulario(diaFormulario)
                end if

                Domingo = PrimerDiaSemanaSQL(diaFormulario)

                antes = "cal_semanal.asp?d=" &  NuevaFecha(PrimerDiaSemanaFormulario(diaFormulario), -7)
                despues = "cal_semanal.asp?d=" &  NuevaFecha(PrimerDiaSemanaFormulario(diaFormulario), 7)

                vDomingo = VinculoEventos(diaFormulario, 0)
                vLunes = VinculoEventos(diaFormulario, 1)
                vMartes = VinculoEventos(diaFormulario, 2)
                vMiercoles = VinculoEventos(diaFormulario, 3)
                vJueves = VinculoEventos(diaFormulario, 4)
                vViernes = VinculoEventos(diaFormulario, 5)
                vSabado = VinculoEventos(diaFormulario, 6)

                nDomingo = NuevoDia(PrimerDiaSemanaFormulario(diaFormulario), 0)
                nLunes = NuevoDia(PrimerDiaSemanaFormulario(diaFormulario), 1)
                nMartes = NuevoDia(PrimerDiaSemanaFormulario(diaFormulario), 2)
                nMiercoles = NuevoDia(PrimerDiaSemanaFormulario(diaFormulario), 3)
                nJueves = NuevoDia(PrimerDiaSemanaFormulario(diaFormulario), 4)
                nViernes = NuevoDia(PrimerDiaSemanaFormulario(diaFormulario), 5)
                nSabado = NuevoDia(PrimerDiaSemanaFormulario(diaFormulario), 6)

                sqlString = "exec cal_Pivot_SemanaDesglosada '" & Request.Cookies("Usuario") & "', '" & Domingo & "'"

                set t = con.execute(sqlString)
                
            '-- Fin: Preparar Datos                
        %>

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                Semana del <%= diaFormulario %> al <%= NuevaFecha(PrimerDiaSemanaFormulario(diaFormulario), 6) %>
            </div>
            
            <div style="flex: 0 0 50%; text-align: right;">
                <button class='form-btn tiny violeta' 
                        type='button' 
                        onclick="irA('<%= antes %>')"  >
                    <<
                    <!-- Anterior (<%= NuevaFecha(PrimerDiaSemanaFormulario(diaFormulario), -7) %>) -->
                </button>

                <button class='form-btn small violeta' 
                        type='button' 
                        onclick="irA('cal_semanal.asp?d=')"  >
                    Hoy
                </button>

                <button class='form-btn tiny violeta' 
                        type='button' 
                        onclick="irA('<%= despues %>')"  >
                    >>
                    <!-- (<%= NuevaFecha(PrimerDiaSemanaFormulario(diaFormulario), 7) %>) -->
                </button>

                &nbsp;

                <!--

                <button class='form-btn tiny azul' 
                        type='button' 
                        onclick="irA('cal_eventos_editar.asp?o=s&f=<%= diaFormulario %>&s=*')" >
                    &nbsp;+&nbsp;
                </button>                   

                &nbsp;
                
                -->

                <button class='form-btn normal verde' 
                        type='button' 
                        onclick="irA('cal_calendario.asp')" >
                    Volver
                </button>                
            </div>
        </div>        

        <div class="main" style="width: 95%;">
            <div class="line">
                <div class="tabla-wrapper">
                    <table class="tabla tabla-violet" id="calSemanal">
                        <thead>
                            <tr>
                                <th class="sticky" style="text-align: center; width:  9%;">Hora</th>
                                <th class="sticky" style="text-align: center; width: 13%;" onclick="irA('<%= vDomingo   %>')">Domingo   <%= nDomingo   %></th>
                                <th class="sticky" style="text-align: center; width: 13%;" onclick="irA('<%= vLunes     %>')">Lunes     <%= nLunes     %></th>
                                <th class="sticky" style="text-align: center; width: 13%;" onclick="irA('<%= vMartes    %>')">Martes    <%= nMartes    %></th>
                                <th class="sticky" style="text-align: center; width: 13%;" onclick="irA('<%= vMiercoles %>')">Miércoles <%= nMiercoles %></th>
                                <th class="sticky" style="text-align: center; width: 13%;" onclick="irA('<%= vJueves    %>')">Jueves    <%= nJueves    %></th>
                                <th class="sticky" style="text-align: center; width: 13%;" onclick="irA('<%= vViernes   %>')">Viernes   <%= nViernes   %></th>
                                <th class="sticky" style="text-align: center; width: 13%;" onclick="irA('<%= vSabado    %>')">Sábado    <%= nSabado    %></th>
                            </tr>
                        </thead>

                        <tbody>
                            <%
                                Do
                                    %>
                                        <tr>
                                            <td style="font-size: 16px; font-family: Arial; font-weight: bold; text-align: center; width: 9%;"><%= left(t("Hora"), 5) %></td>
                                            <td style="width: 13%;"><%= Parse(t("Domingo")) %></td>
                                            <td style="width: 13%;"><%= Parse(t("Lunes")) %></td>
                                            <td style="width: 13%;"><%= Parse(t("Martes")) %></td>
                                            <td style="width: 13%;"><%= Parse(t("Miercoles")) %></td>
                                            <td style="width: 13%;"><%= Parse(t("Jueves")) %></td>
                                            <td style="width: 13%;"><%= Parse(t("Viernes")) %></td>
                                            <td style="width: 13%;"><%= Parse(t("Sabado")) %></td>
                                        </tr>
                                    <%

                                    t.MoveNext
                                Loop Until (t.eof)
                            %>                                                       
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <% t.close: set t = nothing %>

        <script type="text/javascript">
            function iniciar() {
                irAFila('calSemanal', 16);
            }

            function irAFila(idTabla, indiceFila) {
                const tabla = document.getElementById(idTabla);
                if (!tabla) return;

                const tbody = tabla.tBodies[0];
                const filas = tbody.rows;

                if (indiceFila < 0 || indiceFila >= filas.length) {
                    console.log("Índice de fila fuera de rango.");
                    return;
                }

                const fila = filas[indiceFila];
                const contenedor = tabla.closest('.tabla-wrapper') || tabla.parentElement;

                contenedor.scrollTop =
                    fila.offsetTop - contenedor.offsetTop;
            }          

            function abrir(tipo, llave) {
                var vinculo;

                switch (tipo) {
                    case "cal":
                        vinculo = "cal_eventos_editar.asp?o=s&s=" + llave;
                        break;

                    case "pre":
                        vinculo = "../pre/presupuestos/pre_det_tra_editar.asp?llaveCal=" + llave;
                        break;

                    case "con":
                        vinculo = "../cont/cont_editar.asp?con=" + llave;
                        break;

                    default:
                        // opcional: por si viene algo raro
                        break;
                }                

                window.location.href = vinculo;
            }

            function irA(vinculo) {
                window.location.href = vinculo;
            }         
        </script>

        <% con.close: set con = nothing %>
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>