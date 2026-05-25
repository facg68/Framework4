<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 

<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->   

        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0100"
            SysLockOut    
        %>        

        <style>
            td {
                font-size: 14px;
                font-family: Ruda;
                padding: 3px;
            }
        </style>  

        <%
            function LocalMonetarioUsuario()
                dim fcon, f

                set fcon = Server.CreateObject("ADODB.Connection")
                    fcon.Open Application("Conn")

                    set f = fcon.execute("SELECT isnull(usuLocal, 'US') AS usuLocal FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';")
                        LocalMonetarioUsuario = f("usuLocal")           
                    f.close: set f = nothing
                fcon.close: set fcon = nothing
            end function

            function preLocalDestino(presupuesto)
                dim fcon, f

                set fcon = Server.CreateObject("ADODB.Connection")
                    fcon.Open Application("Conn")

                    set f = fcon.execute("SELECT MonedaDestino FROM pre_Presupuesto_Encabezado WHERE Presupuesto = '" & presupuesto & "' AND Usuario = '" & Request.Cookies("Usuario") & "'")
                        preLocalDestino = f("MonedaDestino")
                    f.close: set f = nothing
                fcon.close: set fcon = nothing
            end function

            function HoraLista(Fechatabla)
                dim t, d, m, a, puntero, k

                a = Year(Fechatabla)
                m = RIGHT("00" & Month(Fechatabla), 2)
                d = RIGHT("00" & Day(Fechatabla), 2)

                puntero = -1

                for k = 1 to len(Fechatabla)
                    if mid(Fechatabla, k, 1) = " " then
                        if puntero < 0 then
                            puntero = k
                        end if
                    end if
                next

                t = mid(Fechatabla, puntero + 1, 12)

                HoraLista = d & "/" & m & "/" & a & "<br/>" & t
            end function
        %>                
    </head>

    <body plantilla="normal" reserva="165" onload="Refrescar()">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            dim con, t, sqlString, cbox, c 
            dim Codigo, Anualidad, TipoCuenta, Tipo, Nombre
            dim Monto, Contacto, MensajeDefault, Clase
            dim LocalMonetario, Grupo, Categoria
            dim Repetitiva, DeSistema

            usu = Request.Cookies("usuario")
            c = Request.QueryString("c")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            if c = ""  then 
                Codigo = "CU-000000000"
                Anualidad = ""
                TipoCuenta = "N"
                Nombre = "Nueva Cuenta"
                Tipo = "-"
                Valor = 0.00
                Contacto = NULL
                MensajeDefault = "Mensaje por Defecto"
                Clase = "N"
                LocalMonetario = LocalMonetarioUsuario()
                Grupo = "A"
                Categoria = "normal"
                Repetitiva = 0
                DeSistema = 0

                RegistroNuevo = "1"
                visHistorial = "none"                
            else

                sqlString = "SELECT Usuario, Codigo, Nombre, Categoria, Tipo, Anualidad, Monto, Contacto, " & _
                                " LocalMonetario, MensajeDefault, TipoCuenta, Repetitiva, DeSistema, Grupo, Clase " & _
                            "FROM pre_Cuentas as c " & _
                            "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                            "AND (Codigo = '" & c & "');"

                set t = con.execute(sqlString)

                    Codigo = t("Codigo")
                    Anualidad = t("Anualidad")
                    TipoCuenta = t("TipoCuenta")
                    Nombre = t("Nombre")
                    Tipo = t("Tipo")
                    Valor = t("Monto")  
                    Contacto = t("Contacto")
                    MensajeDefault = t("MensajeDefault")
                    Clase = t("Clase")
                    LocalMonetario = t("LocalMonetario")
                    Grupo = t("Grupo")
                    Categoria = t("Categoria")
                    Repetitiva = t("Repetitiva")
                    DeSistema = t("DeSistema")

                    RegistroNuevo = "0"
                    visHistorial = "block"                

                t.close: set t = nothing
            end if            
        %>          

        <br />

        <form name="form_transaccion" id="form_transaccion" method="post" action="cuentas_grabar.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    <%
                        if RegistroNuevo = 1 then
                            response.write "Nueva Cuenta"
                        else
                            response.write Nombre
                        end if
                    %>                  
                </div>
                
                <div style="flex: 0 0 50%; text-align: right;">
                    <button type="button" 
                            class="form-btn large verde" 
                            onclick="AbrirVinculo('cuentas_contactos.asp?c=<%= Codigo %>')"
                            id="shares" name="shares" style="visibility: hidden;" >
                        Comparticiones
                    </button>    

                    &nbsp;

                    <button type="button" class="form-btn normal azul" onclick="submit()">
                        <% 
                            if RegistroNuevo = 1 then 
                                response.write "Crear"
                            else
                                response.write "Actualizar"
                            end if
                        %>
                    </button>    

                    &nbsp;

                    <button type="button" class="form-btn normal rojo" onclick="AbrirVinculo('lista.asp')">
                        Cancelar
                    </button>                      
                </div>
            </div>        

            <div class="main main-scroll"> 
                <div class="no-ver">
                    <input id="Nuevo"        name="Nuevo"           type="text"     value="<%= RegistroNuevo %>" />
                    <input id="ordenamiento" name="ordenamiento"    type="text"     value="<%= oParan %>" />
                    <input id="cod"          name="cod"             type="text"     value="<%= codigo %>" />
                </div>

                <!-- Campos -->

                    <div class="line">
                        <label class="label normal">Codigo</label>
                        <input class="field tiny" name="codigo" id="codigo" type="text" value="<%= codigo %>"
                            <%
                                if RegistroNuevo <> 1 then
                                    response.write " disabled"
                                end if
                            %>
                        >
                    </div>

                    <div class="line">
                        <label class="label normal">Funcion</label>

                        <select class="field large" name="tipocuenta" id="tipocuenta" OnChange="ActualizarTipoCuenta()" >
                            <option value="N" <% if tipocuenta = "N" then response.write " selected" %>>Suma / Resta</option>
                            <option value="A" <% if tipocuenta = "A" then response.write " selected" %>>Acumulador</option>
                        </select>   
                    </div>

                    <div class="line">
                        <label class="label normal">Nombre</label>
                        <input class="field large" name="Nombre" type="text" value="<%= nombre %>" >
                    </div>

                    <div class="line">
                        <label class="label normal">Tipo</label>

                        <select class="field small" name="tipo" id="tipo" >
                            <option value="+" <% if tipo = "+" then response.write " selected" %>>Credito</option>
                            <option value="-" <% if tipo = "-" then response.write " selected" %>>Debito</option>
                        </select>   
                    </div>

                    <div id="divAnualidad" style="display: block;">
                        <div class="line">
                            <label class="label normal">Anualidad</label>
                            <input class="field tiny" name="anualidad" id="anualidad" type="text" value="<%= anualidad %>" placeholder="dd/mm" >
                        </div>
                    </div>

                    <div class="line">
                        <label class="label normal">Contacto</label>

                        <select class="field large" name="contacto" id="contacto" >
                            <%
                                sqlString = "SELECT Codigo, PrimerNombre + iif(PrimerApellido <> '', ' ' + PrimerApellido, '') AS NombreComtacto " & _
                                            "FROM con_Contactos " & _
                                            "WHERE usuario = '" & usu & "' " & _
                                            "AND visible = '1' " &  _
                                        "ORDER BY NombreComtacto;"

                                set cbox = con.execute(sqlString)

                                response.write "<option value=''"
                                    if Contacto = "" then response.write " selected = 'selected'"
                                response.write ">&nbsp;</option>"

                                if not (cbox.bof or cbox.eof) then
                                    Do
                                        response.write "<option value='" & cbox("Codigo") & "' "
                                            if Contacto = cbox("Codigo") then 
                                                response.write " selected='selected'" 
                                            end if
                                        response.write ">" & cbox("NombreComtacto") & "</option>"

                                    cbox.MoveNext
                                    Loop Until cbox.eof
                                end if

                                cbox.close: set cbox = nothing
                            %>
                        </select>   
                    </div>

                    <div id="divClase" style="display: block;">
                        <div class="line">
                            <label class="label normal">Clase</label>

                            <select class="field normal" name="clase" id="clase" onchange="verificarClase()" >
                                <option value="N" <% if clase = "N" then response.write " selected" %>>Propia</option>
                                <option value="C" <% if clase = "C" then response.write " selected" %>>Compartida</option>
                                <option value="R" <% if clase = "R" then response.write " selected" %>>Rotativa</option>
                            </select>   
                        </div> 
                    </div>  

                    <div class="line">
                        <label class="label normal">Categoria</label>

                        <select class="field normal" name="categoria" id="categoria" >
                            <%
                                sqlString = "SELECT Codigo, Nombre " & _
                                            "FROM dbo.pre_Cuentas_Categorias AS c " & _
                                            "ORDER BY Nombre;"

                                set cbox = con.execute(sqlString)

                                if not (cbox.bof or cbox.eof) then
                                    Do
                                        response.write "<option value='" & cbox("Codigo") & "' "
                                            if categoria = cbox("Codigo") then 
                                                response.write " selected='selected'" 
                                            end if
                                        response.write ">" & cbox("Nombre") & "</option>"

                                        cbox.MoveNext
                                    Loop Until cbox.eof
                                end if

                                cbox.close: set cbox = nothing
                            %>
                        </select>   
                    </div> 

                    <div class="line">
                        <label class="label normal">Grupo</label>

                        <select class="field normal" name="grupo" id="grupo" >
                            <option value="A" <% if grupo = "A" then response.write " selected" %>>Activa</option>
                            <option value="W" <% if grupo = "W" then response.write " selected" %>>En Espera</option>
                            <option value="S" <% if grupo = "S" then response.write " selected" %>>Archivada</option>
                        </select>   
                    </div>      

                    <div class="line">
                        <label class="label normal">Local Monetario</label>

                        <select class="field normal" name="localmonetario" id="localmonetario" >
                            <%
                                sqlString = "SELECT [Local], Simbolo + '  (' + NombreListas + ')' AS NombreLocal " & _
                                                "FROM seg_Cripto_NumParse_Locales " & _
                                            "WHERE [Local] <> 'NUM' " & _
                                            "ORDER BY NombreLocal ASC;"

                                set cbox = con.execute(sqlString)

                                if not (cbox.bof or cbox.eof) then
                                    Do
                                        response.write "<option value='" & cbox("Local") & "' "
                                            if localmonetario = cbox("Local") then 
                                                response.write " selected='selected'" 
                                            end if
                                        response.write ">" & cbox("NombreLocal") & "</option>"

                                        cbox.MoveNext
                                    Loop Until cbox.eof
                                end if

                                cbox.close: set cbox = nothing
                            %>
                        </select>   
                    </div>    

                    <div id="divMonto" style="display: block;">
                        <div class="line">
                            <label class="label normal">Monto</label>
                            <input class="field tiny" name="monto" id="monto" type="text" placeholder="0.00" value="<%= FormatNumber(Valor) %>" />
                        </div>   
                    </div>

                    <div class="line">
                        <label class="label normal">Mensaje</label>
                        <input class="field xxl" name="mensajedefault" id="mensajedefault" type="text"  value="<%= mensajedefault %>"  />
                    </div>  

                    <div class="line">
                        <label class="label normal">Repetición</label>

                        <select class="field large" name="repetitiva" id="repetitiva" >
                            <option value="0" <% if repetitiva = "0" then response.write " selected" %>>Sólo puede aparecer una vez por presupuesto</option>
                            <option value="1" <% if repetitiva = "1" then response.write " selected" %>>Puede repetirse en un presupuesto</option>
                        </select>   
                    </div>

                <!-- Fin de los Campos -->

                <div id="divHistorial" style="display: <%= visHistorial %>;">
                    <div class="line">
                        <label class="label normal" style="align-items: flex-start;">
                            Transacciones Activas:
                        </label>

                        <label class="label full section">
                            <%
                                sqlString = "pa_pre_Cuentas_Historial_Activo '" & usu & "','" & Codigo & "', '" & LocalMonetarioUsuario() & "'"
                                set cbox = con.execute(sqlString)

                                if not (cbox.bof or cbox.eof) then
                                    sw = 1
                                    cuantos = 0
                                    Total = 0.00
                                    TotalAplicado = 0.00

                                    %>
                                        <table class="tabla tabla-green" style="width: 98%; margin: auto;">
                                            <thead>
                                                <tr>
                                                    <th style="text-align: center;">Fecha</th>
                                                    <th style="text-align: center;">Presupuesto</th>
                                                    <th style="text-align: center;">Descripción</th>
                                                    <th style="text-align: center;">Contacto</th>
                                                    <th style="text-align: center;">Monto</th>
                                                </tr>
                                            </thead>

                                            <tbody>
                                                <%
                                                    do 

                                                        sw = -1 * cint(sw)              
                                                        cuantos = cuantos + 1

                                                        Total = Total + cbox("Valor")

                                                        if cbox("Aplicado") = 1 then
                                                            TotalAplicado = TotalAplicado + cbox("Valor")
                                                        end if

                                                        response.write "<tr>"
                                                            response.write "<td style='text-align: center;'>" & HoraLista(cbox("FechaHora")) & "</td>"
                                                            response.write "<td>" & cbox("Nombre") & "</td>"
                                                        
                                                            response.write "<td>" 
                                                                if cbox("Aplicado") = 0 then
                                                                    response.write "(P) " & cbox("Descripcion") 
                                                                else
                                                                    response.write cbox("Descripcion") 
                                                                end if                  
                                                            response.write "</td>"

                                                            response.write "<td>" & cbox("Contacto") & "</td>"
                                                            response.write "<td style='text-align: right;'>" & FormatNumber(cbox("Valor")) & "</td>"
                                                        response.write "</tr>"

                                                        cbox.MoveNext
                                                    Loop until cbox.eof
                                                %> 
                                            </tbody>

                                            <tfoot>
                                                <tr>
                                                    <td colspan="2" style="text-align: left;">
                                                        <button type="button" class="form-btn xxl verde" onclick="AbrirVinculo('cuentas_historial.asp?c=<%= Codigo %>')">
                                                            Transacciones Históricas
                                                        </button>                                                 
                                                    </td>


                                                    <td colspan="3" style="text-align: right; font-weight: normal;">
                                                        Total Actual:&nbsp;&nbsp;<%= FormatNumber(TotalAplicado) %>&nbsp;&nbsp;
                                                        <br />
                                                        Total Presupuestado:&nbsp;&nbsp;<%= FormatNumber(Total) %>&nbsp;&nbsp;                                    
                                                    </td>
                                                </tr>
                                            </tfoot>
                                        </table>     
                                    <%
                                end if

                                cbox.close: set cbox = nothing
                            %>                               
                        </label>
                    </div>   
                </div>                       
            </div>    
        </form>

        <br /><br />

        <script type="text/javascript">            
            function Refrescar() {
                var c = document.getElementById("codigo").value;

                if (c != "CU-000000000") {
                    document.getElementById("codigo").disabled = true;
                };

                ActualizarTipoCuenta();
                verificarClase();
            }

            function verificarClase() {
                var cla = document.getElementById("clase").value;

                if (cla == "N") {
                    document.getElementById("shares").style.visibility = "hidden"; 
                }
                else {
                    document.getElementById("shares").style.visibility = "visible";             
                }
            }

            function ActualizarTipoCuenta() {
                var tipo = document.getElementById("tipocuenta").value;

                if (tipo == "A") {
                    document.getElementById("anualidad").value = "";
                    document.getElementById("divAnualidad").style.display = "none";

                    document.getElementById("monto").value = 0.00;
                    document.getElementById("divMonto").style.display = "none";

                    document.getElementById("clase").value = "N";
                    document.getElementById("divClase").style.display = "none";

                    verificarClase();
                }
                else {
                    document.getElementById("divAnualidad").style.display = "block";
                    document.getElementById("divMonto").style.display = "block";
                    document.getElementById("divClase").style.display = "block";
                }
            }            
                        
            function submit(){
                document.getElementById("form_transaccion").submit(); 
            }

            function AbrirVinculo(vinculo) {
                window.location.href = vinculo;
            }

            mask(document.getElementById('anualidad'), ['99/99']);                
            mask(document.getElementById('codigo'), ['____________']);                
        </script>

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>