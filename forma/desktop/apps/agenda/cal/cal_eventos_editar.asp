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

            dim dia, mes, amo, usu, secuencia, fServer
            dim sqlString, Tipo, tipoForm, fFiltro
            dim aMesActual, aMesFecha, tituloEvento     
            dim Calendario, Titulo, Fecha, FechaFin
            dim TodoElDia, Repeticion, Direccion, Nota
            dim Presupuesto, Monto, DbCr

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")            

            function cargarFecha(secuencia)
                dim cc, tt, sqlString, ff

                sqlString = "SELECT Fecha FROM cal_Eventos WHERE Secuencia = " & secuencia & ";"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                    set tt = cc.execute(sqlString)
                        f = RIGHT("0" & day(tt("Fecha")), 2) & "/" & RIGHT("0" & month(tt("Fecha")), 2) & "/" & year(tt("Fecha"))
                    tt.close: set tt = nothing
                cc.close: set cc = nothing

                cargarFecha = f
            end function

            function calDefault()
                dim cc, tt, sqlString, ff

                sqlString = "SELECT Codigo " & _
                            "FROM cal_Calendarios " & _
                            "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                            "AND (PorDefecto = 1);"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                    set tt = cc.execute(sqlString)
                        calDefault = tt("Codigo")
                    tt.close: set tt = nothing
                cc.close: set cc = nothing            
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

            sub tablaContactosRelacionados(Secuencia)
                dim tcr_con, tcr_cbox, sqlString

                %>
                    <table class="subformulario" style='width: 100%; padding: 0;'>
                <%

                    set tcr_con = Server.CreateObject("ADODB.Connection")
                    tcr_con.open Application("Conn")
                        sqlString = "SELECT Evento, Usuario, Calendario, Contacto, dbo.con_Contactos_NombreContacto(Usuario, Contacto) AS Nombre " & _
                                        "FROM dbo.cal_Eventos_Participantes AS ep " & _
                                    "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _ 
                                        "AND (Evento = " & Secuencia & " ) " & _
                                    "ORDER BY Nombre ASC;"                   

                        set tcr_cbox = tcr_con.execute(sqlString)
                            if not (tcr_cbox.bof or tcr_cbox.eof) then
                                Do
                                    %>
                                        <tr>
                                            <td style="width:90%;">
                                                <input class="field full" type="text" disabled value="<%= tcr_cbox("Nombre") %>" /><br/>
                                            </td>

                                            <td style="width:10%;">
                                                <button type="button" class="form-btn rojo tiny" onclick="borrarContRelacionado('<%= tcr_cbox("Evento") %>','<%= tcr_cbox("Calendario") %>','<%= tcr_cbox("Contacto") %>')">
                                                    <i class="fa fa-trash"></i>
                                                </button>
                                            <td>
                                        </div>
                                    <%

                                    tcr_cbox.MoveNext
                                Loop Until tcr_cbox.eof
                            else
                                response.write "&nbsp;"
                            end if
                        tcr_cbox.close: set tcr_cbox = nothing

                        '
                        ' Añadimos un "formulario" para añadir 
                        '

                        sqlString = "SELECT Codigo, dbo.con_Contactos_NombreContacto(Usuario, Codigo) AS Nombre " & _
                                    "FROM dbo.con_Contactos " & _
                                    "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') "& _
                                    "AND (Visible = 1) " & _
                                    "AND (Codigo NOT IN (" & _
                                        " SELECT Contacto " & _
                                            "FROM dbo.cal_Eventos_Participantes AS ep " & _
                                            "WHERE (Evento = " & secuencia & ") " & _
                                            "AND (Usuario = '" & Request.Cookies("Usuario") & "'))) " & _
                                "ORDER BY Nombre;"

                        set tcr_cbox = tcr_con.execute(sqlString)
                            if not (tcr_cbox.bof or tcr_cbox.eof) then
                                %>
                                    <tr>
                                        <td style="width: 90%;">
                                            <select class="field" style="width: 100%;" name='NuevoContactoRelacionado' id='NuevoContactoRelacionado'>
                                                <option value='*'>- - SELECCIONAR - -</option>
                                                <%
                                                    Do
                                                        response.write "<option value='" & tcr_cbox("Codigo") & "'>" & tcr_cbox("Nombre") & "</option>"
                                                        tcr_cbox.MoveNext
                                                    Loop Until tcr_cbox.eof
                                                %>
                                            </select>
                                        </td>

                                        <td style="width: 10%; text-align: left;">
                                            <button type="button" class="form-btn verde tiny" onclick="NuevoContRelacionado()">
                                                <i class="fa fa-save"></i>
                                            </button>
                                        </td>
                                    </tr>
                                <%
                            end if
                        tcr_cbox.close: set tcr_cbox = nothing
                    tcr_con.close: set tcr_con = nothing    
                response.write "</table>"   
            end sub     

            sub DMA(FechaFormulario)
                Dim p, f

                p = Split(FechaFormulario, "/")
                f = DateSerial(p(2), p(1), p(0))

                dia = Right("0" & Day(f), 2)
                mes = Right("0" & Month(f), 2)
                amo = Year(f)

                fFiltro =  Right("0" & Day(f), 2) & "/" & Right("0" & Month(f), 2) & "/" & Year(f)
            End Sub
        %>   

        <style>
            body { overflow: none; }

            .subformulario tr {
                border-bottom: none !important;
            }            
        </style>         
    </head>

    <body plantilla="normal" reserva="125" onload="iniciar()">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            usu = Request.Cookies("Usuario")
            secuencia = Request.QueryString("s")
            tipoForm = Request.QueryString("o")
            fecha = Request.QueryString("f")
            origen = Request.QueryString("o")

            if (secuencia = "") or (secuencia = NULL) then secuencia = "*"

            if (secuencia = "*") then
                if (fecha = "") or (fecha = NULL) or (fecha = "*") then 
                    fecha = Hoy()
                end if
            else
                fecha = cargarFecha(secuencia)
            end if

            DMA fecha            
            CargarRegistro secuencia, fFiltro
            fserver = FechaServer(fecha)
        %>        

        <br />

        <form id="frm_Eventos" name="frm_Eventos" action="cal_eventos_grabar.asp" method="post"> 
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Editar Evento
                </div>
                
                <div style="flex: 0 0 50%; text-align: right;">
                    <button type="button" class="form-btn verde normal" onclick="GrabarRegistro()">
                        Grabar
                    </button>    

                    &nbsp;

                    <button type="button" class="form-btn rojo normal" onclick="borrarEvento('<%= secuencia %>', '<%= Titulo %>', '<%= dia %>', <%= mes %>, <%= amo %>)">
                        Borrar
                    </button>

                   &nbsp;

                    <button type="button" class="form-btn azul normal" onclick="volver('<%= origen %>')">
                        Volver
                    </button>                         
                </div>
            </div>        

            <div class="main main-scroll">
                <div class="no-ver">
                    <input name="secuencia" id="secuencia"  type="text"   value="<%= secuencia %>">
                    <input name="fFiltro"   id="fFiltro"    type="text"   value="<%= fFiltro %>">
                    <input name="origen"    id="origen"     type="text"   value="<%= origen %>">
                    <input name="dia"       id="dia"        type="text"   value="<%= dia %>">
                    <input name="mes"       id="mes"        type="text"   value="<%= mes %>">
                    <input name="amo"       id="amo"        type="text"   value="<%= amo %>">
                </div>

                <div class="line">
                    <label class="label normal">Calendario</label>
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
                    <label class="label normal">Título</label>
                    <input class="field xxl" id="titulo" name="titulo" placeholder="Descripción del evento" type="text" value="<%= Titulo %>">
                </div>

                <div class="line">
                    <label class="label normal">Desde</label>
                    <input class="field small" 
                           id="Fecha" name="Fecha" 
                           type="text" 
                           placeholder="dd/mm/aaaa hh:mm" 
                           value= "<%= Fecha %>"
                           onchange="refrescarFechas()" >               
                </div>

                <div class="line">
                    <label class="label normal">Hasta</label>
                    <input class="field small" 
                           id="FechaFin" name="FechaFin" 
                           type="text" 
                           placeholder="dd/mm/aaaa hh:mm" 
                           value= "<%= FechaFin %>"
                           onchange="refrescarFechas()" >    
                </div>

                <div class="line">
                    <label class="label normal">Tipo de Evento</label>
                    <select class="field xxl" name="TodoElDia" id="TodoElDia">
                        <option value = "0" <% if TodoElDia = 0 then response.write " selected " %>>El evento sólo aplica a la hora especificada</option>
                        <option value = "1" <% if TodoElDia = 1 then response.write " selected " %>>El evento se extiende todo el día</option>
                    </select>                    
                </div>

                <div class="line">
                    <label class="label normal">Repetición</label>
                    <select class="field large" name="Repeticion" id="Repeticion">
                        <option value = "0" <% if Repeticion = 0 then response.write " selected " %>>El evento no se repite</option>
                        <option value = "1" <% if Repeticion = 1 then response.write " selected " %>>El evento se repite una vez al mes</option>
                        <option value = "2" <% if Repeticion = 2 then response.write " selected " %>>El evento se repite una vez al año</option>
                    </select>                    
                </div>

                <div class="line">
                    <label class="label normal" for="mensaje">Dirección:</label>
                    <textarea class="field xxl" iname="Direccion" id="Direccion" rows="2"><%= Direccion %></textarea>
                </div>

                <div class="line">
                    <label class="label normal" for="mensaje">Nota:</label>
                    <textarea class="field xxl" iname="Nota" id="Nota" rows="2"><%= Nota %></textarea>
                </div>                

                <div class="line">
                    <label class="label normal" for="mensaje">Presupuesto:</label>
                    <select class="field large" name="Presupuesto" id="Presupuesto" onchange="togglePre()">
                        <option value = "0" <% if Presupuesto = 0 then response.write " selected " %>>No es afectado</option>
                        <option value = "1" <% if Presupuesto = 1 then response.write " selected " %>>Se aplica un Monto</option>
                    </select>                    
                </div>

                <div id="DatosPresupuesto" style="display: none;">
                    <div class="line">
                        <label class="label normal" for="mensaje">Monto:</label>
                        <input class="field small" id="Monto" name="Monto" type="number" step="0.01" value="<%= Monto %>" placeholder="0.00">

                        &nbsp;&nbsp;

                        <select class="field small" name="DbCr" id="DbCr">
                            <option value = "0" <% if DbCr = 0 then response.write " selected " %>>Debito</option>
                            <option value = "1" <% if DbCr = 1 then response.write " selected " %>>Credito</option>
                        </select>
                    </div>
                </div>

                <% if Secuencia <> "*" then %>
                    <div class="line">
                        <label class="label normal" for="mensaje">Contactos:</label>
                        <label class="label section" style="width:95%;">
                            <% tablaContactosRelacionados Secuencia %>
                        </label>
                    </div>  
                <% end if %>                    
            </div>    
        </form>

        <br /><br />

        <script type="text/javascript">
            function iniciar() {
                togglePre();
            }

            function GrabarRegistro() {
                document.getElementById("frm_Eventos").submit();          
            }

            function borrarEvento(secuencia, titulo, dia, mes, amo) {
                var confirmacion = confirm("Está seguro de borrar el evento:\n\n'" + titulo + "' ?");

                if (confirmacion) {
                    var vinculo = "cal_eventos_borrar.asp?s=" + secuencia + "&d=" + dia + "&m=" + mes + "&a=" + amo;
                    window.location.href = vinculo;
                }        
            } 

            function togglePre() {
                var pre = document.getElementById("DatosPresupuesto");
                var p = document.getElementById("Presupuesto").value;

                var eM = document.getElementById("Monto");
                var eD = document.getElementById("DbCr");

                if (p == 1) {   
                    eM.alue="0.00";
                    pre.style.display = "block";
                } else {
                    eM.alue="0.00";
                    pre.style.display = "none";
                };
            }
            
            function NuevoContRelacionado() {
                var evento = document.getElementById("secuencia").value;
                var calendario = document.getElementById("Calendario").value;
                var contacto = document.getElementById("NuevoContactoRelacionado").value;

                if (contacto != "*") {
                    var vinculo = "cal_eventos_contactos_grabar.asp?ev=" + evento + "&cal=" + calendario + "&con=" + contacto;
                    window.location.href = vinculo;
                }
            }

            function borrarContRelacionado(evento, calendario, contacto) {
                var confirmacion = confirm("Está seguro de borrar este contacto?");

                if (confirmacion) {
                    var vinculo = "cal_eventos_contactos_borrar.asp?ev=" + evento + "&cal=" + calendario + "&con=" + contacto ;
                    window.location.href = vinculo;
                } else {
                    alert("Proceso Cancelado.");
                }        
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

            function volver(origen) {
                const dia = document.getElementById("dia").value;
                const mes = document.getElementById("mes").value;
                const amo = document.getElementById("amo").value;

                switch (origen) {
                    case "d":
                        window.location.href = `cal_eventos.asp?d=${dia}&m=${mes}&a=${amo}`;
                        break;
                    case "s":
                        window.location.href = "cal_semanal.asp?f=" + fFiltro;
                        window.location.href = `cal_semanal.asp?f=${dia}/${mes}/${amo}`;
                        break;
                    case "m":
                        window.location.href = `cal_calendario.asp?m=${mes}&a=${amo}`;
                        break;
                }                
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

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>