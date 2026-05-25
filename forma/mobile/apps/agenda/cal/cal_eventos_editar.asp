<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Editar Evento" %>
        <title><%= PageTitle %></title>

        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0020"
            SysLockOut

            dim dia, mes, amo, usu, secuencia, fServer
            dim sqlString, Tipo, tipoForm, fFiltro
            dim aMesActual, aMesFecha, tituloEvento     
            dim Calendario, Titulo, Fecha, FechaFin
            dim TodoElDia, Repeticion, Direccion, Nota
            dim Presupuesto, Monto, DbCr

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn") 

            ' Funciones y Procedimientos -------------------------------------------------------------------------------------

                function cargarFecha(secuencia)
                    dim tt, sqlString, ff

                    sqlString = "SELECT FORMAT(Fecha, 'dd/MM/yyyy') AS FechaFormulario " & _
                                "FROM dbo.cal_Eventos WHERE Secuencia = " & secuencia & ";"

                    set tt = con.execute(sqlString)
                        f = tt("FechaFormulario")
                    tt.close: set tt = nothing

                    cargarFecha = f
                end function  

                function calDefault()
                    dim tt, sqlString, ff

                    sqlString = "SELECT Codigo " & _
                                "FROM cal_Calendarios " & _
                                "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                "AND (PorDefecto = 1);"

                    set tt = con.execute(sqlString)
                        calDefault = tt("Codigo")
                    tt.close: set tt = nothing
                end function   

                function Hoy()
                    Dim p, f

                    p = Split(Date(), "/")
                    f = DateSerial(CInt(p(2)), CInt(p(0)), CInt(p(1)))
                    Hoy = Right("0" & Day(f), 2) & "/" & Right("0" & Month(f), 2) & "/" & Year(f)
                end function

                function FechaServer(FechaFormulario)
                    Dim d, m, a

                    d = left(FechaFormulario, 2)
                    m = mid(FechaFormulario, 4, 2)
                    a = mid(FechaFormulario, 7, 4)

                    f = DateSerial(a, m, d)
                    FechaServer = Year(f) & "-" & Right("0" & Month(f), 2) & "-" & Right("0" & Day(f), 2) 
                end function      

                function FechaFormulario(FechaServer)
                    dim dia, mes, amo, horas, minutos

                    dia = Day(FechaServer)
                    mes = Month(FechaServer)
                    amo = Year(FechaServer)

                    horas = Hour(FechaServer)
                    minutos = Minute(FechaServer)

                    FechaFormulario = right("00" & dia, 2) & "/" & right("00" & mes, 2) & "/" & amo 
                    FechaFormulario = FechaFormulario & " " & right("00" & horas, 2) & ":" & right("00" & minutos, 2)  
                end function  

                Function AjustarFechaHora(fechaTexto, Horas)
                    Dim fBase, fAhora, fFinal

                    ' Convertir la fecha dd/MM/aaaa a Date
                    fBase = DateSerial( _
                                CInt(Mid(fechaTexto, 7, 4)), _
                                CInt(Mid(fechaTexto, 4, 2)), _
                                CInt(Mid(fechaTexto, 1, 2)) _
                            )

                    ' Tomar hora y minutos actuales
                    fAhora = TimeSerial(Hour(Now()), Minute(Now()), 0)

                    ' Unir fecha + hora actual
                    fFinal = fBase + fAhora

                    ' Sumar o restar horas
                    fFinal = DateAdd("h", Horas, fFinal)

                    ' Devolver como texto dd/MM/aaaa HH:mm
                    AjustarFechaHora = _
                        Right("0" & Day(fFinal), 2) & "/" & _
                        Right("0" & Month(fFinal), 2) & "/" & _
                        Year(fFinal) & " " & _
                        Right("0" & Hour(fFinal), 2) & ":" & _
                        Right("0" & Minute(fFinal), 2)

                End Function   
                
                Sub CargarRegistro(Secuencia, fechaFiltro)
                    dim tt

                    if (secuencia = "*") then
                        Calendario = calDefault()
                        Titulo = ""
                        Fecha = AjustarFechaHora(fechaFiltro, 0)
                        FechaFin = AjustarFechaHora(fechaFiltro, 1)
                        TodoElDia = 0
                        Repeticion = 0
                        Direccion = ""
                        Nota = ""
                        Presupuesto = 0        
                        Monto = 0.00   
                        DbCr = 0                    
                    else
                        set tt = con.execute("SELECT * FROM cal_Eventos WHERE Secuencia = " & secuencia & ";")
                            Calendario = tt("Calendario")
                            Titulo = tt("Titulo")
                            Fecha = FechaFormulario(tt("Fecha"))
                            FechaFin = FechaFormulario(tt("FechaFin"))
                            TodoElDia = tt("TodoElDia")
                            Repeticion = tt("Repeticion")
                            Direccion = tt("Direccion")
                            Nota = tt("Nota")
                            Presupuesto = tt("Presupuesto")
                            Monto = tt("Monto") 
                            DbCr = tt("DbCr")
                        tt.close: set tt = nothing
                    end if            
                End Sub   

                sub DMA(FechaFormulario)
                    Dim p

                    p = Split(FechaFormulario, "-")
                    fFiltro =  p(2) & "/" & p(1) & "/" & p(0)
                End Sub                                                                            

            '-----------------------------------------------------------------------------------------------------------------
        %>           
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <%
            usu = Request.Cookies("Usuario")
            secuencia = Request.QueryString("s")
            fecha = Request.QueryString("f")

            if fecha <> "" then DMA fecha   
            if (secuencia = "") or (secuencia = NULL) then secuencia = "*"

            if (secuencia = "*") then
                if (fecha = "") or (fecha = NULL) or (fecha = "*") then 
                    fecha = Hoy()
                end if
            else
                fecha = cargarFecha(secuencia)
            end if

            CargarRegistro secuencia, fFiltro
            fserver = FechaServer(fecha)
        %>  

        <div class="page-title-bar">
            <div class="title"><%= PageTitle %></div>
        </div>

        <form id="frm_Eventos" name="frm_Eventos" action="cal_eventos_grabar.asp" method="post"> 
            <div class="no-ver">
                <input name="secuencia" id="secuencia"  type="text"   value="<%= secuencia %>">
                <input name="fFiltro"   id="fFiltro"    type="text"   value="<%= fFiltro %>">
                <input name="dia"       id="dia"        type="text"   value="<%= dia %>">
                <input name="mes"       id="mes"        type="text"   value="<%= mes %>">
                <input name="amo"       id="amo"        type="text"   value="<%= amo %>">
                <input name="TodoElDia" id="TodoElDia"  type="text"   value="0">
            </div>
                    
            <main>
                <br />

                <div class="contenedor">
                     <div class="line">
                        <label>Título</label>
                        <input id="titulo" name="titulo" placeholder="Descripción del evento" type="text" value="<%= Titulo %>">
                    </div>

                    <div class="line">
                        <label>Desde</label>
                        <input id="Fecha" name="Fecha" type="text" placeholder="dd/mm/aaaa hh:mm" value= "<%= Fecha %>" onchange="refrescarFechas()" >               
                    </div>

                    <div class="line">
                        <label>Hasta</label>
                        <input id="FechaFin" name="FechaFin" type="text" placeholder="dd/mm/aaaa hh:mm" value= "<%= FechaFin %>" onchange="refrescarFechas()" >    
                    </div>

                    <div class="line">
                        <label>Repetición</label>
                        <select name="Repeticion" id="Repeticion">
                            <option value = "0" <% if Repeticion = 0 then response.write " selected " %>>El evento no se repite</option>
                            <option value = "1" <% if Repeticion = 1 then response.write " selected " %>>El evento se repite una vez al mes</option>
                            <option value = "2" <% if Repeticion = 2 then response.write " selected " %>>El evento se repite una vez al año</option>
                        </select>                    
                    </div>

                   <div class="line">
                        <label>Calendario</label>
                        <select class="field large" name="Calendario" id="Calendario" required>
                            <%
                                sqlString = "SELECT Codigo, Nombre FROM cal_Calendarios WHERE Usuario = '" & Request.Cookies("Usuario") & "' ORDER BY Nombre ASC;"
                                set tt = con.execute(sqlString)

                                Do
                                    response.write "<option value = '" & tt("Codigo") & "'" 
                                        if tt("Codigo") = Calendario then
                                            response.write " selected "
                                        end if
                                    response.write ">" & tt("Nombre") & "</option>"
                                    
                                    tt.MoveNext
                                Loop until (tt.eof)

                                tt.close: set t = nothing
                            %>
                        </select>
                    </div>                    

                    <div class="line">
                        <label>Dirección:</label>
                        <textarea name="Direccion" id="Direccion" rows="2"><%= Direccion %></textarea>
                    </div>

                    <div class="line">
                        <label for="mensaje">Nota:</label>
                        <textarea name="Nota" id="Nota" rows="2"><%= Nota %></textarea>
                    </div>   
                <div>

                <br />
            </main>
        </form>

        <footer class="footer-contextual">
            <button class="footer-button" type="button" aria-label="Home" onclick="irA('/forma/mobile')">
                <i class="fa-solid fa-house"></i>
            </button>
                    
            <button class="footer-button" type="button" aria-label="Volver" onclick="volver()">
                <i class="fas fa-undo-alt"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Grabar" onclick="grabar()">
                <i class="fas fa-save"></i>
            </button>
        </footer>

        <script>
             function volver() {
                history.back();
            }

            function irA(vinculo) {
                window.location.href = vinculo;
            }

            function grabar() {
                if (validarCampos()) {
                    document.getElementById("frm_Eventos").submit();
                }
            }

            function validarCampos() {
                const titulo   = document.getElementById("titulo").value.trim();
                const fecha    = document.getElementById("Fecha").value.trim();
                const fechaFin = document.getElementById("FechaFin").value.trim();

                if (!titulo || !fecha || !fechaFin) {
                    let mensaje = "Faltan los siguientes campos:<br><br>";

                    if (!titulo)   mensaje += "• Título<br>";
                    if (!fecha)    mensaje += "• Fecha desde<br>";
                    if (!fechaFin) mensaje += "• Fecha hasta<br>";

                    Swal.fire({
                        icon: "warning",
                        title: "Datos incompletos",
                        html: mensaje,
                        confirmButtonText: "Entendido",
                        confirmButtonColor: "#007224",
                        background: "#f9f9f9"
                    });

                    return false;
                }

                return true;
            }

            function refrescarFechas() {
                const inputFecha = document.getElementById("Fecha");
                const inputFechaFin = document.getElementById("FechaFin");

                let fecha = parsearFecha(inputFecha.value);

                // Fecha debe ser válida, si no, fecha = Now()
                if (!fecha) {
                    fecha = new Date();
                }

                // Actualizar campos dia, mes, amo desde Fecha
                const dia = String(fecha.getDate()).padStart(2, "0");
                const mes = String(fecha.getMonth() + 1).padStart(2, "0");
                const amo = fecha.getFullYear();

                document.getElementById("fFiltro").value = `${dia}/${mes}/${amo}`;
                document.getElementById("dia").value = dia;
                document.getElementById("mes").value = mes;
                document.getElementById("amo").value = amo;

                // Normalizar Fecha
                inputFecha.value = formatearFecha(fecha);

                // FechaFin válida, si no FechaFin = Fecha + 1 hora
                let fechaFin = parsearFecha(inputFechaFin.value);

                if (!fechaFin) {
                    fechaFin = new Date(fecha.getTime() + 60 * 60 * 1000);
                }

                // FechaFin no puede ser menor que Fecha, si no FechaFin = Fecha + 1 hora
                if (fechaFin < fecha) {
                    fechaFin = new Date(fecha.getTime() + 60 * 60 * 1000);
                }

                // Normalizar FechaFin
                inputFechaFin.value = formatearFecha(fechaFin);
            }            

            function parsearFecha(valor) {
                if (!valor) return null;

                const partes = valor.split(" ");
                if (partes.length !== 2) return null;

                const [fecha, hora] = partes;
                const f = fecha.split("/");
                const h = hora.split(":");

                if (f.length !== 3 || h.length !== 2) return null;

                const dia = Number(f[0]);
                const mes = Number(f[1]) - 1;
                const amo = Number(f[2]);
                const horas = Number(h[0]);
                const minutos = Number(h[1]);

                const date = new Date(amo, mes, dia, horas, minutos);

                if (
                    date.getFullYear() !== amo ||
                    date.getMonth() !== mes ||
                    date.getDate() !== dia ||
                    date.getHours() !== horas ||
                    date.getMinutes() !== minutos
                ) {
                    return null;
                }

                return date;
            }

            function formatearFecha(date) {
                const d = String(date.getDate()).padStart(2, "0");
                const m = String(date.getMonth() + 1).padStart(2, "0");
                const y = date.getFullYear();
                const h = String(date.getHours()).padStart(2, "0");
                const min = String(date.getMinutes()).padStart(2, "0");

                return `${d}/${m}/${y} ${h}:${min}`;
            }            

            mask(document.getElementById('Fecha'),    ['99/99/9999 99:99']);
            mask(document.getElementById('FechaFin'), ['99/99/9999 99:99']);            
        </script>

        <% con.close: set con = nothing %>
        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>