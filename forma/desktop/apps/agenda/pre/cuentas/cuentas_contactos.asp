<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <title>Editar Cuenta</title>    
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->             
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0100"
            SysLockOut    


            function LocalMonetarioUsuario()
                dim fcon, f

                set fcon = Server.CreateObject("ADODB.Connection")
                fcon.Open Application("Conn")
                    set f = fcon.execute("SELECT isnull(usuLocal, 'US') AS usuLocal FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';")
                        LocalMonetarioUsuario = f("usuLocal")            
                    f.close: set f = nothing
                fcon.close: set fcon = nothing
            end function

            function TipoComparticion(Cuenta)
                dim fcon, f, ssql

                ssql = "SELECT Clase " & _
                       "FROM pre_Cuentas " & _
                       "WHERE Usuario = '" & Request.Cookies("Usuario") & "' " & _
                       "AND Codigo = '" & Cuenta & "';"

                set fcon = Server.CreateObject("ADODB.Connection")
                fcon.Open Application("Conn")
                    set f = fcon.execute(ssql)
                        TipoComparticion = f("Clase")                
                    f.close: set f = nothing
                fcon.close: set fcon = nothing
            end function      

            function CantidadCompartida(Cuenta)
                dim fcon, f, ssql

                ssql = "SELECT Monto " & _
                       "FROM pre_Cuentas " & _
                       "WHERE Usuario = '" & Request.Cookies("Usuario") & "' " & _
                       "AND Codigo = '" & Cuenta & "';"

                set fcon = Server.CreateObject("ADODB.Connection")
                fcon.Open Application("Conn")
                    set f = fcon.execute(ssql)
                        if (f.bof or f.eof) then
                            CantidadCompartida = 0.00
                        else
                            CantidadCompartida = CDbl(f("Monto"))
                        end if
                    f.close: set f = nothing
                fcon.close: set fcon = nothing
            end function  

            function TotalUsuarios(Cuenta)
                dim fcon, f, ssql

                ssql = "SELECT COUNT(*) AS Cuantos " & _
                       "FROM pre_Cuentas_Comparticiones " & _
                       "WHERE Usuario = '" & Request.Cookies("Usuario") & "' " & _
                       "AND Cuenta = '" & Cuenta & "';"

                set fcon = Server.CreateObject("ADODB.Connection")
                fcon.Open Application("Conn")
                    set f = fcon.execute(ssql)
                        TotalUsuarios = f("Cuantos")                
                    f.close: set f = nothing
                fcon.close: set fcon = nothing
            end function        

            function MontoUsuario(Cuenta)     
                Tipo = TipoComparticion(Cuenta)
                Monto = (CantidadCompartida(Cuenta) * 1.00)
                Usuarios = (TotalUsuarios(Cuenta) * 1) + 1

                Select Case Tipo
                    Case "R"
                        MontoUsuario = Monto
                    Case "C"
                        if Usuarios = 0 then
                            MontoUsuario = Monto
                        else
                            MontoUsuario = (Monto / Usuarios)
                        end if
                    Case Else
                        MontoUsuario = 0.00
                End Select
            end function

            function FechaFormulario(FechaServer)
                dim t, d, m, a, puntero, k

                if FechaServer <> "" then
                    a = Year(FechaServer)
                    m = RIGHT("00" & Month(FechaServer), 2)
                    d = RIGHT("00" & Day(FechaServer), 2)

                    puntero = -1

                    for k = 1 to len(FechaServer)
                        if mid(FechaServer, k, 1) = " " then
                            if puntero < 0 then
                                puntero = k
                            end if
                        end if
                    next

                    t = mid(FechaServer, puntero + 1, 12)

                    FechaFormulario = d & "/" & m & "/" & a & " " & t
                else
                    FechaFormulario = ""
                end if
            end function

            function NombreCuenta(Codigo)
                dim cc, tt, ssql

                ssql = "SELECT Nombre FROM pre_Cuentas " & _
                       "WHERE Usuario = '" & Request.Cookies("Usuario") & "' " & _
                       "AND Codigo = '" & Codigo & "';"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.execute(ssql)

                if not (tt.bof or tt.eof) then
                    NombreCuenta = tt("Nombre")
                else
                    NombreCuenta = ""
                end if

                tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function
        %>

        <style>
            .fila {
                display: flex;
                align-items: center;
                gap: 10px;
            }

            /* Columna 1: texto fijo */
            .col1 {
                white-space: nowrap;
                font-weight: bold;
                min-width: fit-content;
            }

            /* Columna 2: se estira todo lo que pueda */
            .col2 {
                flex: 1;
            }

            /* Columna 3: 10% del ancho total */
            .col3 {
                flex: 0 0 10%;
                min-width: 80px; /* opcional, evita que colapse */
            }

            /* Columna 4: botón de 60x60 */
            .col4 {
                flex: 0 0 60px;
                display: flex;
                justify-content: center;
                align-items: center;
            }

            .col4 {
                width: 60px;
                height: 60px;
            }
        </style>    
    </head>

    <body plantilla="tabla" reserva="200">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            dim con, t, sqlString, cbox, cuantos
            dim Secuencia, Usuario, Cuenta, Contacto, MontoCompartido, UltimaFechaAplicada, Puntero

            Usuario = Request.Cookies("usuario")
            Cuenta = Request.QueryString("c")
            cuantos = 0

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            sqlString = "SELECT cu.Secuencia, cu.Usuario, cu.Cuenta, cu.Contacto, co.Nombre, cu.MontoCompartido, cu.UltimaFechaAplicada, cu.Puntero " & _
                        "FROM dbo.pre_Cuentas_Comparticiones AS cu " & _
                        "INNER JOIN dbo.con_Contactos_ListaContactos AS co " & _
                        "ON cu.Usuario = co.Usuario " & _
                        "AND cu.Contacto = co.Codigo " & _
                        "WHERE (cu.Usuario = '" & Request.Cookies("Usuario") & "') " & _
                        "AND (cu.Cuenta = '" & Cuenta & "');"

            set t = con.execute(sqlString)
        %>  

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 90%; text-align: left; font-size: 20px; color: rgb(50, 50, 50);">
                <input style="width: 100%; background-color: rgb(235, 235, 235);" 
                       readonly 
                       value="Contactos en la Cuenta Compartida <%= NombreCuenta(Cuenta) %>">
            </div>
            
            <div style="flex: 0 0 10%; text-align: right;">
                <button type="button" class="form-btn normal azul" onClick="vinculo('cuentas_editar.asp?c=<%= Cuenta %>')">Finalizar</button>
            </div>
        </div>        

        <div class="main" style="width: 95%;">
            <div class="line">
                <div class="tabla-wrapper">
                    <table class="tabla tabla-green">
                        <thead>
                            <tr>
                                <th class="sticky" style="width: 45%; text-align: center;">Contacto</th>
                                <th class="sticky" style="width: 15%; text-align: center;">Monto</th>
                                <th class="sticky" style="width: 15%; text-align: center;">Ult. Pago</th>
                                <th class="sticky" style="width: 25%; text-align: center;"></th>
                            </tr>
                        </thead>

                        <tbody>                    
                    <%
                        if not (t.bof or t.eof) then
                            sw = -1

                            Do
                                cuantos = cuantos + 1

                                response.write "<tr>"
                                    response.write "<td>" & t("Nombre") & "</td>"
                                    response.write "<td style='text-align: right;'>" & FormatNumber(t("MontoCompartido")) & "</td>"
                                    response.write "<td style='text-align: center;'>" & FechaFormulario(t("UltimaFechaAplicada")) & "</td>"

                                    response.write "<td style='text-align: center;'>"
                                        response.write "<a href='cuentas_contactos_puntero.asp?s=" & t("Secuencia") & "'>"
                                            response.write "<button class='form-btn small azul'>" 
                                                response.write "<i class=' fa fa-arrow-left' title='Mover Puntero'></i>"
                                            response.write "</button>"
                                        response.write "</a>"

                                        response.write "&nbsp;"

                                        response.write "<a href='cuentas_contactos_borrar.asp?s=" & t("Secuencia") & "'>"
                                            response.write "<button class='form-btn small rojo'>" 
                                                response.write "<i class=' fa fa-trash' title='borrar contacto compartido'></i>"
                                            response.write "</button>"
                                        response.write "</a>"                                      
                                    response.write "</td>"
                                response.write "</tr>"                        

                                t.MoveNext
                            Loop Until t.eof
                        end if
                    %>
                        </tbody>

                        <tfoot>
                            <tr>
                                <td class="sticky" style="text-align: center;" colspan="4">
                                    <%
                                        if cuantos = 0 then
                                            response.write "&nbsp;&nbsp;Aún no se ha compartido esta cuenta con ningún contacto."
                                        else
                                            response.write "&nbsp;&nbsp;Se está compartiendo con " & cuantos & " contactos."
                                        end if
                                    %>
                                </td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>

            <!--
                Ahora añadimos un formulario para añadir un contacto nuevo
            -->            

            <form name="form_transaccion" id="form_transaccion" method="post" action="cuentas_contactos_grabar.asp">
                <div class="fila">
                    <div class="col1">Nuevo Contacto</div>

                    <div class="col2">
                        <select class="field normal" name="contacto" id="contacto" style="width: 100%;">
                            <%
                                sqlString = "SELECT DISTINCT c.Codigo, c.Nombre " & _
                                                       "FROM dbo.con_Contactos_ListaContactos AS c " & _
                                                 "INNER JOIN dbo.con_Contactos_ConCategs AS ccc " & _
                                                         "ON c.Usuario = ccc.Usuario " & _
                                                        "AND c.Codigo = ccc.Codigo " & _
                                                      "WHERE (c.Usuario = '" & Usuario & "') " & _
                                                        "AND (ccc.Tipo = 'PE') " & _
                                                        "AND (c.Activo = 1) " & _
                                                        "AND (c.Codigo NOT IN (" & _
                                                                                "SELECT Contacto " & _
                                                                                  "FROM dbo.pre_Cuentas_Comparticiones AS pcc " & _
                                                                                 "WHERE (Usuario = '" & Usuario & "') " & _
                                                                                   "AND (Cuenta = '" & Cuenta & "'))) " & _
                                                  "ORDER BY c.Nombre;"

                                set cbox = con.execute(sqlString)

                                if not (cbox.bof or cbox.eof) then
                                    Do
                                        response.write "<option value='" & cbox("Codigo") & "'>" & cbox("Nombre") & "</option>"
                                        cbox.MoveNext
                                    Loop Until cbox.eof
                                end if

                                cbox.close: set cbox = nothing
                            %>
                        </select>   
                    </div>

                    <div class="col3">
                        <input class="field normal" style="width: 100%; text-align: right;" id="mm" name="mm" type="text" value="<%= FormatNumber(MontoUsuario(Cuenta)) %>" disabled/>
                    </div>

                    <div class="col4">
                        <button class="form-btn tiny verde" type="submit">
                            <i class="fa fa-save fa-xl" title="Añadir"></i>
                        </button>  
                    </div>
                </div>

                <div class="no-ver">
                    <input id="cuenta" name="cuenta" type="text" value="<%= Cuenta %>" />
                    <input id="monto"  name="monto"  type="text" value="<%= FormatNumber(MontoUsuario(Cuenta)) %>" />
                </div>
            </form>
        </div>
  
        <br /><br />   

        <script type="text/javascript">
            function submit(){
                document.getElementById("anualidad").disabled = false;
                document.getElementById("monto").disabled = false;
                document.getElementById("clase").disabled = false;

                document.getElementById("form_transaccion").submit(); 
            }

            function vinculo(direccion) {
                window.location.href = direccion;
            }

            mask(document.getElementById('anualidad'), ['99/99']);               
        </script>

        <% con.close: set con = nothing %> 
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>