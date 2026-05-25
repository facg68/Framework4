<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->    
        <% PageTitle = "Calendario Diario" %>
        <title><%= PageTitle %></title>

        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0020"
            SysLockOut


            dim con, t, sqlString, PrimerDia, primeraVez
            dim dia, mes, amo, diaSQL, divID, divIDContador
            dim div_contactoCorreo, div_contactoNombre, div_contactoTelefono, div_presupuesto

            divIDContador = 0

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

            Function NombreFecha(fechaTexto)
                Dim p, f, wDay

                p = Split(fechaTexto, "/")
                f = DateSerial(CInt(p(2)), CInt(p(1)), CInt(p(0)))
                wDay = WeekDay(f)

                NombreFecha = NombreDia(wDay) & " " & CInt(p(0)) & " de " & NombreMes(CInt(p(1))) & " de " & CInt(p(2))
            End function

            Sub CargarDatosDeContacto(contacto)
                dim tt, sqlString

                sqlString = "SELECT TOP (1) c.PrimerNombre + ' ' + c.PrimerApellido AS Nombre, " & _
                                          " c.CorreoElectronico AS Correo, t.Telefono " & _
                            "FROM dbo.con_Contactos AS c LEFT OUTER JOIN dbo.con_Contactos_Telefonos AS t " & _
                            "ON c.Usuario = t.Usuario AND c.Codigo = t.Codigo " & _
                            "WHERE (c.Usuario = '" & Request.Cookies("Usuario") & "') " & _
                            "AND (c.Codigo = '" & contacto & "') " & _
                            "ORDER BY t.Tipo;"

                set tt = con.execute(sqlString)
                    if not (tt.bof or tt.eof) then
                        div_contactoCorreo   = tt("Correo")
                        div_contactoNombre   = tt("Nombre")
                        div_contactoTelefono = tt("Telefono")
                    end if
                tt.close: set tt = nothing
            End Sub

            Sub CargarDatosDePresupuesto(transaccion)
                dim tt, sqlString

                sqlString = "SELECT Presupuesto " & _
                            "FROM dbo.pre_Presupuesto_Detalles " & _
                            "WHERE (Llave = " & transaccion & ")  " & _
                            "AND (Usuario = '" & Request.Cookies("Usuario") & "');"

                set tt = con.execute(sqlString)
                    if not (tt.bof or tt.eof) then
                        div_presupuesto = tt("Presupuesto")
                    end if
                tt.close: set tt = nothing            
            End Sub
       
            Function Parse(texto)
                Dim partes, i, bloque, mensaje, meta, datos, llave, tipo
                Dim salida, pIni, pFin, primero, listaData, divID

                salida = ""
                Parse = ""

                If IsNull(texto) Then Exit Function

                texto = CStr(texto)
                texto = Replace(texto, "<br />", "<br/>", 1, -1, vbTextCompare)
                texto = Replace(texto, "<br>", "<br/>", 1, -1, vbTextCompare)

                If Trim(texto) <> "" Then texto = texto & "<br/>"

                partes = Split(texto, "<br/>")
                primero = 1

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

                                ' Completar Datos -----------------------------------------------------
                                    divID = "calDet_" & RIGHT("000000" & divIDContador, 5)
                                    div_contactoCorreo   = ""
                                    div_contactoNombre   = ""
                                    div_contactoTelefono = ""  
                                    div_presupuesto = ""

                                    if tipo = "con" then CargarDatosDeContacto llave
                                    if tipo = "pre" then CargarDatosDePresupuesto llave

                                    listaData = "id='" & divID & "' " & _
                                                "data-tipo='" & tipo & "' " & _
                                                "data-llave='" & llave & "' " & _
                                                "data-con_nombre='"   & div_contactoNombre   & "' " & _
                                                "data-con_telefono='" & div_contactoTelefono & "' " & _
                                                "data-con_correo='"   & div_contactoCorreo   & "' " & _
                                                "data-presupuesto='"  & div_presupuesto & "'"
                                '----------------------------------------------------------------------

                                if primero = 1 then
                                    primero = 0
                                    salida = salida & "<div class='evento' " & listaData & " >" & mensaje & "</div>"
                                else
                                    salida =  salida & "<br /><div class='evento' " & listaData & " >" & mensaje & "</div>"
                                end if

                                divIDContador = divIDContador + 1
                            End If
                        End If
                    End If
                Next

                Parse = salida
            End Function
        %>    

        <style>
            .cal-evento {
                font-family: Ruda;
                display: flex;
                align-items: flex-start;
                padding: 12px 14px;
                text-decoration: none;
                border-bottom: 1px solid #eee;
                gap: 12px;
            }

            .cal-evento:nth-child(odd) {
                background-color: #fafafa;
            }

            .cal-evento:nth-child(even) {
                background-color: #f3fbf4;
            }

            .cal-hora {
                width: 70px;
                flex-shrink: 0;
                font-family: 'Ruda Bold';
                font-size: 1.2rem;
                text-align: center;
                color: #333; 
            }

            .cal-contenido {
                flex: 1;
                font-size: 1.2rem;
                line-height: 1.4rem;
                color: #333;
            }        
        </style>          
    </head>

    <body plus= ".evento : plus_calDetalles; ">
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     
        <%
            dia = Request.QueryString("d")
            mes = Request.QueryString("m")
            amo = Request.QueryString("a")

            if dia = "" then
                dia = Request.Cookies("cal_dia")
                mes = Request.Cookies("cal_mes")
                amo = Request.Cookies("cal_amo")
            else
                Response.Cookies("cal_dia") = dia
                Response.Cookies("cal_mes") = mes
                Response.Cookies("cal_amo") = amo
            end if

            if dia = "" then 
                diaFormulario = Hoy()

                dia = Day(date())
                mes = Month(date())
                amo = Year(date())


            else
                diaFormulario = Fecha(dia, mes, amo)
            end if       

            diaSQL = FechaSQL(dia, mes, amo)  
            sqlString = "exec cal_Pivot_DiaDesglosado '" & Request.Cookies("Usuario") & "', '" & diaSQL & "'"
            set t = con.execute(sqlString)    
        %>        

        <div class="page-title-bar">
            <div class="title"><%= NombreFecha(diaFormulario)  %></div>
        </div>

        <main>
            <%
                if not (t.bof or t.eof) then
                    Do
                        %>
                            <div class="cal-evento" data-hora="<%= left(t("Hora"),5) %>">
                                <a class="cal-hora">
                                    <%= left(t("Hora"), 5) %>
                                </a>

                                <div class="cal-contenido">
                                    <%= Parse(t("Descripciones")) %>
                                </div>
                            </div>                            
                         <%
                        t.MoveNext
                    Loop Until t.eof
                end if
            %>
        </main>

        <footer class="footer-contextual">
            <button class="footer-button" type="button" aria-label="Home" onclick="irA('/forma/mobile')">
                <i class="fa-solid fa-house"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Volver" onclick="volver()">
                <i class="fa-solid fa-rotate-left"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Ayer" onclick="ayer('<%= diaSQL %>')">
                <i class="fa-solid fa-backward"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Nuevo Evento" onclick="nuevo('<%= diaSQL %>')">
                <i class="fa-solid fa-plus"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Mañana" onclick="manana('<%= diaSQL %>')">
                <i class="fa-solid fa-forward"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Hoy" onclick="hoy('<%= diaSQL %>')">
                <i class="fa-solid fa-bullseye"></i>
            </button>                                                

            <button class="footer-button" type="button" aria-label="Calendario Anual" onclick="anual()">
                <i class="fa-solid fa-calendar"></i>
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
                var vinculo = "cal_eventos_editar.asp?f=" + fechaSQL + "&s=*";
                window.location.href = vinculo;
            }
            
            function ayer(fechaSQL) {
                const partes = fechaSQL.split("-");
                let amo = parseInt(partes[0], 10);
                let mes = parseInt(partes[1], 10) - 1; // JS meses 0-11
                let dia = parseInt(partes[2], 10);

                const fecha = new Date(amo, mes, dia);
                fecha.setDate(fecha.getDate() - 1);

                const nuevoDia = String(fecha.getDate()).padStart(2, "0");
                const nuevoMes = String(fecha.getMonth() + 1).padStart(2, "0");
                const nuevoAmo = String(fecha.getFullYear());

                var vinculo = "cal_calendario.asp?d=" + nuevoDia + "&m=" +  nuevoMes + "&a=" + nuevoAmo;
                window.location.href = vinculo;
            }          
            
            function manana(fechaSQL) {
                const partes = fechaSQL.split("-");
                let amo = parseInt(partes[0], 10);
                let mes = parseInt(partes[1], 10) - 1; // Meses en JS: 0-11
                let dia = parseInt(partes[2], 10);

                const fecha = new Date(amo, mes, dia);
                fecha.setDate(fecha.getDate() + 1);

                const nuevoDia = String(fecha.getDate()).padStart(2, "0");
                const nuevoMes = String(fecha.getMonth() + 1).padStart(2, "0");
                const nuevoAmo = String(fecha.getFullYear());

                var vinculo = "cal_calendario.asp?d=" + nuevoDia + "&m=" +  nuevoMes + "&a=" + nuevoAmo;
                window.location.href = vinculo;
            }   
            
            function hoy() {
                const fecha = new Date();

                const dia = String(fecha.getDate()).padStart(2, '0');
                const mes = String(fecha.getMonth() + 1).padStart(2, '0'); // Enero = 0
                const amo = fecha.getFullYear();

                var vinculo = "cal_calendario.asp?d=" + dia + "&m=" +  mes + "&a=" + amo;
                window.location.href = vinculo;
            }   
            
            function anual() {
                var vinculo = "cal_anual.asp?a=<%= Request.Cookies("cal_amo") %>";
                window.location.href = vinculo;
            }            
            
            document.addEventListener("DOMContentLoaded", () => {
                const contenedor = document.querySelector("main");
                const objetivo = document.querySelector('[data-hora="06:30"]');

                if (contenedor && objetivo) {
                    contenedor.scrollTop = objetivo.offsetTop;
                }
            });            
        </script>

        <% 
            t.close: set t = nothing 
            con.close: set con = nothing
        %>

        <!-- #include virtual = "/forma/plus/telCalDetalles.plus" -->     
        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>