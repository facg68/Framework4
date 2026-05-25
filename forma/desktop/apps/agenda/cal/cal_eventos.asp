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
                font-family: Arial; 
                font-size: 16px;
                text-align: left;
                padding: 0px;
            }

            .hora {
                font-size: 16px; 
                font-family: Arial; 
                font-weight: bold; 
                text-align: center
            }
        </style>  

        <%
            dim con, t, sqlString, PrimerDia, primeraVez
            dim dia, mes, amo, diaSQL

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

            function NombreDia(Dia)
                dim cc, tt, sqlString

                sqlString = "SELECT Valor, Nombre FROM seg_cripto_Secuencias WHERE (Tipo = 'S') AND (Valor = '" & Dia & "');"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                    set tt = cc.execute(sqlString)
                        NombreDia = tt("Nombre")
                    tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function               

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

            Function Fecha(d, m, a)
                dim f

                f = DateSerial(CInt(a), CInt(m), CInt(d))

                Fecha = Right("0" & Day(f), 2) & "/" & _
                        Right("0" & Month(f), 2) & "/" & _
                        Year(f)
            end function

            Function FechaSQL(d, m, a)
                dim f

                f = DateSerial(CInt(a), CInt(m), CInt(d))

                FechaSQL = Year(f) & "-" & _
                           Right("0" & Month(f), 2) & "-" & _
                           Right("0" & Day(f), 2)
            end function            

            Function Hoy()
                dim f, d, m, a

                d = day(date())
                m = month(date())
                a = year(date())

                f = DateSerial(CInt(a), CInt(m), CInt(d))

                Hoy = Right("0" & Day(f), 2) & "/" & _
                      Right("0" & Month(f), 2) & "/" & _
                      Year(f)            
            end Function

            Function CrearVinculo(FechaFormulario)
                Dim p
                p = Split(FechaFormulario, "/")
                CrearVinculo = "cal_eventos.asp?d=" & p(0) & "&m=" & p(1) & "&a=" & p(2)
            End Function

            Function NombreFecha(fechaTexto)
                Dim p, f, wDay

                p = Split(fechaTexto, "/")
                f = DateSerial(CInt(p(2)), CInt(p(1)), CInt(p(0)))
                wDay = WeekDay(f)

                NombreFecha = NombreDia(wDay) & " " & CInt(p(0)) & " de " & NombreMes(CInt(p(1))) & " de " & CInt(p(2))
            End function
        %>                       
    </head>

    <body plantilla="tabla" reserva="125" onload="iniciar()">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->
        <%
            origen = Request.QueryString("o")
            
            dia = Request.QueryString("d")
            mes = Request.QueryString("m")
            amo = Request.QueryString("a")

            if dia = "" then 
                diaFormulario = Hoy()
            else
                diaFormulario = Fecha(dia, mes, amo)
            end if            
       
            diaSQL = FechaSQL(dia, mes, amo)  
            antes = CrearVinculo(NuevaFecha(diaFormulario, -1))
            despues = CrearVinculo(NuevaFecha(diaFormulario, 1))

            sqlString = "exec cal_Pivot_DiaDesglosado '" & Request.Cookies("Usuario") & "', '" & diaSQL & "'"

            set t = con.execute(sqlString)            
        %>

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                <%= NombreFecha(diaFormulario) %>
            </div>
            
            <div style="flex: 0 0 50%; text-align: right;">
                <button class='form-btn tiny violeta'            
                        type='button' 
                        onclick="irA('<%= antes %>')"  >
                    &nbsp;<<&nbsp;
                    <!-- (<%= NuevaFecha(diaFormulario, -1) %>)-->
                </button>

                <button class='form-btn tiny violeta'           
                        type='button' 
                        onclick="irA('<%= despues %>')"  >
                        &nbsp;>>&nbsp;
                        <!-- (<%= NuevaFecha(diaFormulario, +1) %>) -->
                </button>

                &nbsp;

                <button class='form-btn tiny azul' 
                        type='button' 
                        onclick="irA('cal_eventos_editar.asp?o=d&f=<%= diaFormulario %>&d=<%= Dia %>&m=<%= Mes %>&a=<%= Amo %>&s=*')" >
                    &nbsp;+&nbsp;
                </button>

                &nbsp;

                <button class='form-btn normal verde' 
                        type='button' 
                        onclick="irA('cal_calendario.asp')" >
                    Volver
                </button>                
            </div>
        </div>        

        <div class="main" style="width: 95%;"> <!-- ES UNA TABLA -- NO NECESITA MAIN-SCROLL -->
            <div class="line">
                <div class="tabla-wrapper">
                    <table class="tabla tabla-violet" id="calDiario">
                        <thead>
                            <tr>
                                <th class="sticky" style="text-align: center; width: 10%;">Hora</th>
                                <th class="sticky" style="text-align: left; width: 90%;">Eventos</th>
                            </tr>
                        </thead>

                        <tbody>
                            <%
                                Do
                                    %>
                                        <tr>
                                            <td class="hora" style="width: 10%; text-align: center;"><%= left(t("Hora"), 5) %></td>
                                            <td style="width: 90%;"><%= Parse(t("Descripciones")) %></td>
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
                irAFila('calDiario', 16);
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
                        vinculo = "cal_eventos_editar.asp?o=d&d=<%= Dia %>&m=<%= Mes %>&a=<%= Amo %>&s=" + llave;
                        break;
                                            
                    case "pre":
                        vinculo = "../pre/presupuestos/pre_det_tra_editar.asp?llaveCal=" + llave;
                        break;

                    case "con":
                        vinculo = "../cont/cont_editar.asp?con=" + llave;
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