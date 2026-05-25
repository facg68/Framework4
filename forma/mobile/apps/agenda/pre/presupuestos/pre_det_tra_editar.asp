<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Editar Transacción" %>
        <title><%= PageTitle %></title>

        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0090"
            SysLockOut


            dim con, t, p, sqlString, cbox, llave, nuevo, usu, pre, llaveCal, dia

            dim Fecha, Hora, CuentaOrigen, CuentaDestino, MontoOrigen
            dim MontoDestino, MonedaOrigen, MonedaDestino, Descripcion
            dim Contacto, Aplicado, HoraTemp, Estatus, TipoPre, Archivado
            dim redirectVer, redirectTipo, redirectEstatus, redirectOrdenamiento

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            llave = Request.QueryString("registro")
            usu = Request.Cookies("Usuario")

            ' Funciones y Procedimientos ---------------------------------------------------------------   
                function EstatusPresupuesto()
                    dim ta

                    set ta = con.Execute("SELECT Estatus from pre_Presupuesto_Encabezado " & _
                                         " WHERE (Usuario = '" & usu & "') " & _
                                        " AND (Presupuesto = '" & pre &  "');")

                        if not (ta.bof or ta.eof) then
                            EstatusPresupuesto = ta("Estatus")
                        else
                            EstatusPresupuesto = ""
                        end if
                    ta.close: set ta = nothing
                end Function

                function NombrePresupuesto()
                    dim ta

                    set ta = con.Execute("SELECT nombre from pre_Presupuesto_Encabezado " & _
                                        " WHERE (Usuario = '" & usu & "') " & _
                                        " AND (Presupuesto = '" & pre & "');")

                        if not (ta.bof or ta.eof) then
                            NombrePresupuesto = ta("nombre")
                        else
                            NombrePresupuesto = ""
                        end if
                    ta.close: set ta = nothing
                end Function

                Function FechaForm(FechaDB)
                    dim a, m, d

                    FechaForm = ""

                    if not isnull(FechaDB) then
                        d = RIGHT("00" & day(FechaDB), 2)
                        m = RIGHT("00" & month(FechaDB), 2)
                        a = year(FechaDB)

                        FechaForm = d & "/" & m & "/" & a
                    end if
                end function

                Function HoraForm(HoraDB)
                    dim h, m

                    HoraForm = ""

                    if not isnull(HoraDB) then
                        h = LEFT(HoraDB, 2)
                        m = RIGHT(HoraDB, 2)

                        HoraForm = h & ":" & m
                    end if
                end function

                Function MonedaUsuario()
                    dim f, sqlString

                    sqlString = "SELECT usuLocal FROM seg_Usuarios WHERE usuCodigo = '" & usu & "';"

                    set f = con.execute(sqlString)
                        if (f.eof or f.bof) then
                            MonedaUsuario = "US"
                        else
                            if isnull(f("usuLocal")) then
                                MonedaUsuario = "US"
                            else
                                MonedaUsuario = f("usuLocal")
                            end if
                        end if
                    f.close: set f = nothing
                end Function

                function limpiar(cadena)
                    limpiar = cadena
                    
                    limpiar = Replace(limpiar, "&#11013;", "<")
                    limpiar = Replace(limpiar, "&#11157;", ">")
                end function   
            
                function preLocalOrigen()
                    dim f

                    set f = con.execute("SELECT MonedaOrigen " & _
                                        " FROM pre_Presupuesto_Encabezado " & _
                                        " WHERE Presupuesto = '" & pre &  "' " & _
                                        " AND Usuario = '" & usu & "';")

                        preLocalOrigen = f("MonedaOrigen")
                    f.close: set f = nothing
                end function

                function preLocalDestino()
                    dim f

                    set f = con.execute("SELECT MonedaDestino FROM pre_Presupuesto_Encabezado " & _
                                        " WHERE Presupuesto = '" & pre &  "' " & _
                                        " AND Usuario = '" & usu & "';")

                        preLocalDestino = f("MonedaDestino")
                    f.close: set f = nothing
                end function

                Function HoraVbs()
                    dim h, m
                    
                    h = RIGHT("00" & Hour(Time()), 2)
                    m = RIGHT("00" & Minute(Time()), 2)

                    HoraVbs = h & ":" & m
                end function  

                Function MultiPrecio()
                    dim ta

                    set ta = con.Execute("SELECT multiprecio from pre_Presupuesto_Encabezado " & _
                                        " WHERE (Usuario = '" & usu & "') " & _
                                        " AND (Presupuesto = '" & pre &  "');")

                        if not (ta.bof or ta.eof) then
                            MultiPrecio = ta("multiprecio")
                        else
                            MultiPrecio = 0
                        end if
                    ta.close: set ta = nothing
                end function    

                function sysDateTimeOffset()
                    dim ta

                    set ta = con.Execute("SELECT dbo.sysDateTimeOffset() AS FechaHoraSistema;")
                        sysDateTimeOffset = ta("FechaHoraSistema")
                    ta.close: set ta = nothing
                end function    

                function sysDateTimeOffset_FechaForm()
                    dim d, m, a, f
                    f = sysDateTimeOffset()

                    d = mid(f, 9, 2)
                    m = mid(f, 6, 2)
                    a = left(f, 4)

                    sysDateTimeOffset_FechaForm = right("00" & d, 2) & "/" & right("00" & m, 2) & "/" & a
                end function    

                function sysDateTimeOffset_HoraForm()
                    dim h, m, f
                    f = sysDateTimeOffset()

                    h = right("00" & mid(f, 12, 2), 2)
                    m = right("00" & mid(f, 15, 2), 2)

                    sysDateTimeOffset_HoraForm = h & ":" & m
                end function    

                function Presupuesto_Llave()
                    dim ta, sqlString

                    sqlString = "SELECT Presupuesto " & _
                                "FROM dbo.pre_Presupuesto_Detalles " & _
                                "WHERE (Llave = " & Llave & ");"

                    set ta = con.Execute(sqlString)
                        if not (ta.bof or ta.eof) then
                            Presupuesto_Llave = ta("Presupuesto")
                        else
                            Presupuesto_Llave = NULL
                        end if
                    ta.close: set ta = nothing
                end Function     
                
                function tipoPresupuesto()
                    dim ta, sqlString

                    sqlString = "SELECT Tipo " & _
                                "FROM dbo.pre_Presupuesto_Encabezado " & _
                                "WHERE (Presupuesto = '" & pre &  "');"

                    set ta = con.Execute(sqlString)
                        if not (ta.bof or ta.eof) then
                            tipoPresupuesto = ta("Tipo")
                        else
                            tipoPresupuesto = NULL
                        end if
                    ta.close: set ta = nothing
                end Function 
            '-------------------------------------------------------------------------------------------
        %>   

        <style>
            .swal-white-clean {
                border-radius: 18px;
                box-shadow: 0 20px 45px rgba(0,0,0,0.15);
            }   
            
            .line {
                margin-bottom: 18px;
            }

            .fecha-control,
            .hora-control {
                display: flex;
                align-items: center;
                gap: 12px;
            }

            .fecha-control input,
            .hora-control input {
                flex: 1;
                min-width: 0;
                text-align: center;
            }
           
            .fecha-control button,
            .hora-control button {
                width: 44px;
                height: 44px;
                background: #2e7d32;   /* ajusta al verde exacto de tu app */
                color: #ffffff;
                border: none;
                border-radius: 10px;
                font-size: 18px;
                font-weight: 600;
                display: flex;
                align-items: center;
                justify-content: center;
                cursor: pointer;
                transition: transform 0.08s ease, filter 0.15s ease;
            }

            .fecha-control button:active,
            .hora-control button:active {
                transform: scale(0.95);
                filter: brightness(0.9);
            }     
            
            .cuenta-control {
                display: flex;
                align-items: center;
                gap: 12px;
            }

            .cuenta-control select {
                flex: 1;
                min-width: 0;
            }

            .cuenta-control button {
                width: 44px;
                height: 44px;
                background: #2e7d32; /* tu verde sexy */
                color: #ffffff;
                border: none;
                border-radius: 10px;
                font-size: 18px;
                font-weight: 600;
                display: flex;
                align-items: center;
                justify-content: center;
                cursor: pointer;
                transition: transform 0.08s ease, filter 0.15s ease;
            }

            .cuenta-control button:active {
                transform: scale(0.95);
                filter: brightness(0.9);
            }            
        </style>       
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->   
        <%
            if llave = 0  then 
                nuevo = 1
                pre = Request.QueryString("pre")
                Estatus = EstatusPresupuesto()
                TipoPre = tipoPresupuesto()
            else
                nuevo = 0    
                pre = Presupuesto_Llave()
                Estatus = EstatusPresupuesto()
                TipoPre = tipoPresupuesto()                    
            end if
        %>          

        <div class="page-title-bar">
            <div class="title"><%= PageTitle %></div>
        </div>

        <%
            if nuevo = 1 then 
                Fecha = sysDateTimeOffset_FechaForm
                Hora = sysDateTimeOffset_HoraForm
                CuentaOrigen = "PRE-000"
                CuentaDestino = "SYS-000"
                MontoOrigen = 0.00
                MontoDestino = 0.00
                MontoCambio = 0.00
                MonedaOrigen = preLocalOrigen()
                MonedaDestino = preLocalDestino()
                Descripcion = ""
                Contacto = ""
                Aplicado = 0
                Archivado = 0                
            else
                sqlString = "SELECT d.Llave, d.Presupuesto, d.Usuario, d.Fecha, d.Hora, d.CuentaOrigen, d.MontoOrigen, e.MonedaOrigen, " & _
                                " d.Descripcion, d.CuentaDestino, d.MontoDestino, e.MonedaDestino, d.MontoCambio, d.Aplicado, d.Archivado, " & _
                                " d.Contacto, d.Nota, d.Incremento " & _
                            "FROM dbo.pre_Presupuesto_Detalles AS d " & _
                        "INNER JOIN dbo.pre_Presupuesto_Encabezado AS e " & _
                                "ON d.Presupuesto = e.Presupuesto " & _
                            "AND d.Usuario = e.Usuario " & _
                            "WHERE Llave = " & llave & ";"

                set t = con.execute(sqlString)
                    HoraTemp = RIGHT("0000" & t("Hora"), 4)

                    Fecha = FechaForm(t("Fecha"))
                    Hora = HoraForm(HoraTemp)
                    CuentaOrigen = t("CuentaOrigen")
                    CuentaDestino = t("CuentaDestino")
                    MontoOrigen = t("MontoOrigen")
                    MontoDestino = t("MontoDestino")
                    MontoCambio = t("MontoCambio")
                    MonedaOrigen = t("MonedaOrigen")
                    MonedaDestino = t("MonedaDestino")
                    Descripcion = t("Descripcion")
                    Contacto = t("Contacto")
                    Aplicado = t("Aplicado")
                    Archivado = t("Archivado")
                t.close: set t = nothing
            end if  
        %>

        <form name="form_transaccion" id="form_transaccion" method="post" action="pre_det_grabar.asp">
            <div class="no-ver">
                <input id="usuario"         name="usuario"          type="text" value="<%= usu %>"           />
                <input id="presupuesto"     name="presupuesto"      type="text" value="<%= pre %>"           /> 
                <input id="d"               name="d"                type="text" value="<%= dia %>"           /> 
                <input id="MonedaOrigen"    name="MonedaOrigen"     type="text" value="<%= MonedaOrigen %>"  />
                <input id="MonedaDestino"   name="MonedaDestino"    type="text" value="<%= MonedaDestino %>" />
                <input id="Nuevo"           name="Nuevo"            type="text" value="<%= Nuevo %>"         />
                <input id="Llave"           name="Llave"            type="text" value="<%= Llave %>"         />
            </div>  

            <main>
                <br />

                <div class="contenedor">
                    <div class="line">
                        <label>Descripcion</label>
                        <input id="Descripcion" name="Descripcion" type="text" placeholder="Nueva Transaccion" value="<%= limpiar(Descripcion) %>" />
                    </div>

                    <div class="line">
                        <label>Cuenta Origen</label>

                        <div class="cuenta-control">
                            <select name="CuentaOrigen" id="CuentaOrigen">
                                <%
                                    sqlString ="pre_Tra_CuentaOrigen '" & usu & "','" & pre & "', '" & CuentaOrigen & "'"

                                    set cbox = con.execute(sqlString)
                                        if not (cbox.bof or cbox.eof) then
                                            Do
                                                response.write "<option value='" & cbox("Codigo") & "' "
                                                    if CuentaOrigen = cbox("Codigo") then 
                                                        response.write " selected='selected'" 
                                                    end if
                                                response.write ">" & cbox("Nombre") & "</option>"

                                                cbox.MoveNext
                                            Loop Until cbox.eof
                                        end if
                                    cbox.close: set cbox = nothing
                                %>
                            </select>

                            <button type="button" onClick="swapCuentas()"> ▼ </button>
                        </div>
                    </div>      

                    <div class="line">
                        <label>Cuenta Destino</label>
                        <label>
                            <select name="CuentaDestino" id="CuentaDestino">
                                <%
                                    sqlString ="pre_Tra_CuentaDestino '" & usu & "','" & pre & "', '" & CuentaDestino & "'"                

                                    set cbox = con.execute(sqlString)
                                        if not (cbox.bof or cbox.eof) then
                                            Do
                                                response.write "<option value='" & cbox("Codigo") & "' "
                                                    if CuentaDestino = cbox("Codigo") then 
                                                        response.write " selected='selected'" 
                                                    end if
                                                response.write ">" & cbox("Nombre") & "</option>"

                                                cbox.MoveNext
                                            Loop Until cbox.eof
                                        end if
                                    cbox.close: set cbox = nothing
                                %>
                            </select>
                        </label>
                    </div>

                    <div class="line">
                        <label>Monto Origen</label>
                        <label>
                            <input id="Monto" name="Monto" type="text" value="<%
                                if MontoOrigen >= 0 then
                                    response.write formatNumber(MontoOrigen)
                                else
                                    response.write FormatNumber( (-1 * MontoOrigen) )
                                end if
                            %>" OnChange="cMoneda1('<%= MonedaOrigen %>','<%= MonedaDestino %>', 1);" />

                            &nbsp;&nbsp;

                            <select name="MonedaOrigenDisplay" id="MonedaOrigenDisplay" disabled >
                                <%
                                    sqlString = "SELECT [Local], Simbolo + ' (' + NombreListas + ')' AS NombreLocal " & _
                                                  "FROM seg_Cripto_NumParse_Locales " & _
                                                "WHERE [Local] <> 'NUM' " & _
                                                "ORDER BY NombreLocal ASC;"

                                    set cbox = con.execute(sqlString)
                                        if not (cbox.bof or cbox.eof) then
                                            Do
                                                response.write "<option value='" & cbox("Local") & "' "
                                                    if MonedaOrigen = cbox("Local") then 
                                                        response.write " selected" 
                                                    end if
                                                response.write ">" & limpiar(cbox("NombreLocal")) & "</option>"

                                                cbox.MoveNext
                                            Loop Until cbox.eof
                                        end if
                                    cbox.close: set cbox = nothing
                                %>
                            </select>                             
                        </label>
                    </div>

                    <div class="<%
                                    if MonedaOrigen = MonedaDestino then
                                        response.write "no-ver"
                                    else
                                        response.write "line"
                                    end if
                                %>">
                        <label>Monto Destino</label>
                        <label>
                            <input id="txtMontoCambio" name="txtMontoCambio" type="text" value="<%
                                if MontoCambio => 0 then
                                    response.write formatNumber(MontoCambio)
                                else
                                    response.write FormatNumber( (-1 * MontoCambio) )
                                end if
                            %>" OnChange="cMoneda2('<%= MonedaDestino %>','<%= MonedaOrigen %>', 2);" />

                            &nbsp;

                            <select name="MonedaDestinoDisplay" id="MonedaDestinoDisplay" disabled>
                                <%
                                    sqlString = "SELECT [Local], Simbolo + ' (' + NombreListas + ')' AS NombreLocal " & _
                                                  "FROM seg_Cripto_NumParse_Locales " & _
                                                "WHERE [Local] <> 'NUM' " & _
                                                "ORDER BY NombreLocal ASC;"

                                    set cbox = con.execute(sqlString)
                                        if not (cbox.bof or cbox.eof) then
                                            Do
                                                response.write "<option value='" & cbox("Local") & "' "
                                                    if MonedaDestino = cbox("Local") then 
                                                        response.write " selected='selected'" 
                                                    end if
                                                response.write ">" & limpiar(cbox("NombreLocal")) & "</option>"

                                                cbox.MoveNext
                                            Loop Until cbox.eof
                                        end if
                                    cbox.close: set cbox = nothing
                                %>
                            </select>                            
                        </label>
                    </div>

                    <div class="line">
                        <label>Fecha</label>

                        <div class="fecha-control">
                            <button type="button" onclick="fechaMenos()">◀</button>
                            <input id="txt_fecha" name="txt_fecha" type="text"
                                value="<%= fecha %>"
                                placeholder="dd/mm/aaaa"
                                onchange="FechaValida();" />
                            <button type="button" onclick="fechaMas()">▶</button>
                        </div>
                    </div>

                    <div class="line">
                        <label>Hora</label>

                        <div class="hora-control">
                            <button type="button" onclick="horaMenos()">◀</button>
                            <input id="txt_hora"
                                name="txt_hora"
                                type="text"
                                value="<%= hora %>"
                                placeholder="hh:mm"
                                onchange="HoraValida();" />
                            <button type="button" onclick="horaMas()">▶</button>
                        </div>
                    </div>

                    <div class="line">
                        <label>Contacto</label>
                        <select name="Contacto" id="Contacto">
                            <%
                                sqlString = "SELECT Codigo, PrimerNombre + iif(PrimerApellido <> '', ' ' + PrimerApellido, '') AS NombreComtacto " & _
                                                "FROM con_Contactos " & _
                                            "WHERE usuario = '" & usu & "' " & _
                                            "AND (visible = 1) " &  _
                                            "AND (estatus = 1) " &  _
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
                                        response.write ">" & limpiar(cbox("NombreComtacto")) & "</option>"

                                        cbox.MoveNext
                                    Loop Until cbox.eof
                                end if

                                cbox.close: set cbox = nothing
                            %>
                        </select>                              
                    </div>                                        

                    <div class="line">
                        <label>Aplicar</label>
                        <select name="Aplicado" id="Aplicado" >
                            <option value="0" <% if Aplicado = 0 then response.write " selected" %>>No Aplicado</option>
                            <option value="1" <% if Aplicado = 1 then response.write " selected" %>>Aplicado</option>
                        </select>                           
                    </div>

                    <div class="line">
                        <label>Acción</label>
                        <select name="preSiguiente" id="preSiguiente" class="foot">
                            <option value="*"><%
                                    if TipoPre = "M" then
                                        response.write "Transacción del Modelo Actual" 
                                    else
                                        response.write "Transacción del Presupuesto Actual" 
                                    end if
                                %>
                            </option>

                            <%
                                If Nuevo = 0 then
                                    '
                                    ' Las transacciones NUEVAS no se pueden copiar o mover...
                                    ' Sólo se pueden crear
                                    '
                                    if (Estatus = 1) AND (Aplicado = 0) AND (Archivado = 0) then
                                        '
                                        ' Solo los Presupuestos o Modelos ACTIVOS pueden mover transacciones...
                                        ' Y Sólo a Presupuestos o Modelos Activos...
                                        '
                                        sqlString = "SELECT Presupuesto, Nombre " & _
                                                        "FROM pre_Presupuesto_Encabezado " & _
                                                        "WHERE (Usuario = '" & usu & "') " & _
                                                        "AND (Tipo = '" & TipoPre & "') " & _ 
                                                        "AND (Estatus = 1) " & _ 
                                                        "AND (Presupuesto <> '" & pre & "') " & _
                                                    "ORDER BY Presupuesto;"

                                        '
                                        ' Mover Transacciones
                                        '

                                        set cbox = con.execute(sqlString)
                                            if not (cbox.bof or cbox.eof) then
                                                Do
                                                    lblContador = "Mover Transaccion a&nbsp;[&nbsp;"
                                                    response.write "<option value='M@" & cbox("Presupuesto") & "'>" & lblContador & cbox("Nombre") & "&nbsp;]</option>"
                                                    
                                                    cbox.MoveNext
                                                Loop Until cbox.eof
                                            end if
                                        cbox.close: set cbox = nothing

                                        '
                                        ' Copiar Transacciones
                                        '                            

                                        set cbox = con.execute(sqlString)
                                            if not (cbox.bof or cbox.eof) then
                                                Do
                                                    lblContador = "Copiar Transaccion a&nbsp;[&nbsp;"
                                                    response.write "<option value='C@" & cbox("Presupuesto") & "'>" & lblContador & cbox("Nombre") & "&nbsp;]</option>"
                                                    
                                                    cbox.MoveNext
                                                Loop Until cbox.eof
                                            end if
                                        cbox.close: set cbox = nothing

                                    else
                                        response.write "&nbsp;"
                                    end if
                                end if
                            %>
                        </select>  
                    </div>

                    <br />
                <div>
            </main>
        </form>

        <footer class="footer-contextual">
            <button class="footer-button" type="button" aria-label="Home" onclick="irA('/forma/mobile')">
                <i class="fa-solid fa-house"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Volver" onclick="volver()">
                <i class="fas fa-undo-alt"></i>
            </button>

            <% if (Estatus > 0) AND (Archivado = 0) then %>
                <button class="footer-button" type="button" aria-label="Grabar" onclick="grabar()">
                    <i class="fas fa-save"></i>
                </button>
            <% end if %>
        </footer>

        <script>
            function volver() {
                history.back();
            }

            function irA(vinculo) {
                window.location.href = vinculo;
            }

            function grabar() {
                if ((HoraValida() == 0) && (FechaValida() == 0) && (CantidadValida() == 0)) {
                    document.getElementById("form_transaccion").submit();
                } else {
                    swalert("Algunos de los valores entrados no son correctos. Por favor, verifique y vuelva a intentarlo.");
                }
            }    
            
            function swapCuentas() {
                var c;

                c = document.getElementById("CuentaOrigen").value

                document.getElementById("CuentaOrigen").value = document.getElementById("CuentaDestino").value;
                document.getElementById("CuentaDestino").value = c;
            }

            function fechaMas() {
                var f

                f = sumarUnDia(document.getElementById("txt_fecha").value);
                document.getElementById("txt_fecha").value = f;
            }    

            function fechaMenos() {
                var f

                f = restarUnDia(document.getElementById("txt_fecha").value);
                document.getElementById("txt_fecha").value = f;
            }    
            
            function horaMas() {
                var h

                h = sumar30Minutos(document.getElementById("txt_hora").value);
                document.getElementById("txt_hora").value = h;
            }    

            function horaMenos() {
                var h

                h = restar30Minutos(document.getElementById("txt_hora").value);
                document.getElementById("txt_hora").value = h;
            }       
            
            function FechaValida() {
                var sw = 1;
                var valor = document.getElementById("txt_fecha").value;
                var amo, mes;

                var dd = valor.substring(0, 2);
                var mm = valor.substring(3, 5);        

                dia = parseInt(dd);
                mes = parseInt(mm);

                if ((dia >= 1 && dia <= 31) && (mes >= 1 && mes <= 12)) {
                    return 0;
                } else {
                    swalert("El valor de la Fecha no es válido. Por favor verifique.");
                    document.getElementById("txt_fecha").value = "<%= fecha %>";
                };
            };

            function HoraValida() {
                var sw = 1;        
                var valor = document.getElementById("txt_hora").value;
                var hora, min;

                var hh = valor.substring(0, 2);
                var mm = valor.substring(3, 5);

                hora = parseInt(hh);
                min = parseInt(mm);

                if ((hora >= 0 && hora <= 23) && (min >= 0 && min <= 59)) {
                    return 0;
                } else {
                    swalert("El valor de la Hora no es válido. Por favor verifique.");  
                    document.getElementById("txt_hora").value = "<%= hora %>";        
                };
            };

            function CantidadValida() {
                var c1, c2;
                var mOrigen = document.getElementById("Monto").value;
                var mDestino = document.getElementById("txtMontoCambio").value;

                mOrigen = mOrigen.replace(/,/g, "");
                mDestino = mDestino.replace(/,/g, "");
                
                c1 = !isNaN(mOrigen); 
                c2 = !isNaN(mDestino); 

                if (c1 && c2) { return 0; }
                else { return 1; }        
            };

            function cMoneda1(desde, hasta, m) {
                var c1;
                var mOrigen = document.getElementById("Monto").value;

                mOrigen = mOrigen.replace(/,/g, "");
                c1 = !isNaN(mOrigen); 

                if (c1) {
                    CambiarMoneda(desde, hasta, m);
                } else {
                    swalert("El valor del Monto es incorrecto. Verifique y vuelva a intentarlo.");

                    document.getElementById("Monto").value = "<%
                        if MontoOrigen >= 0 then
                            response.write formatNumber(MontoOrigen)
                        else
                            response.write FormatNumber( (-1 * MontoOrigen) )
                        end if          
                    %>";
                }        
            };

            function cMoneda2(desde, hasta, m) {
                var c2;
                var mDestino = document.getElementById("txtMontoCambio").value;

                mDestino = mDestino.replace(/,/g, "");
                c2 = !isNaN(mDestino); 

                if (c2) {
                    CambiarMoneda(desde, hasta, m);
                } else {
                    swalert("El valor del Monto Destino es incorrecto. Verifique y vuelva a intentarlo.");

                    document.getElementById("txtMontoCambio").value = "<%
                        if MontoCambio >= 0 then
                            response.write formatNumber(MontoCambio)
                        else
                            response.write FormatNumber( (-1 * MontoCambio) )
                        end if          
                    %>";              
                }        
            };

            function CambiarMoneda(desde, hasta, m) {
                var donde = 0;
                var k = 0;
                var monto = 0.00;
                var monto2 = 0.00;

                if (m == 1) {
                    monto = document.getElementById("Monto").value;

                    if (monto < 0) {
                        monto = (-1 * monto);

                        swalert("El Valor SIEMPRE debe ser mayor o igual a cero.");
                        document.getElementById("Monto").value = monto;
                    };          
                };

                if (m == 2) {
                    monto = document.getElementById("txtMontoCambio").value

                    if (monto < 0) {
                        monto = (-1 * monto);

                        swalert("El Valor SIEMPRE debe ser mayor o igual a cero.");
                        document.getElementById("txtMontoCambio").value = monto;
                    };          
                };

                var formatter = new Intl.NumberFormat('en-US', {
                    style: 'decimal',
                    currency: 'USD',
                });   

                var locales = [<%
                    set tloc = con.execute("SELECT Local, Simbolo, Formula " & _
                                            "FROM seg_Cripto_NumParse_Locales " & _
                                            "WHERE Local <> 'NUM' " & _
                                            "ORDER BY Local ASC;")

                    if not (tloc.bof or tloc.eof) then
                        response.write "'*'"
                        
                        do
                            response.write ", "
                            response.write "'" & tloc("local") & "'"
                            tloc.MoveNext
                        loop until (tloc.eof)
                    end if

                    tloc.close: set tloc = nothing
                %>];

                var simbolos = [<%
                    set tsim = con.execute("SELECT Local, Simbolo, Formula " & _
                                            "FROM seg_Cripto_NumParse_Locales " & _
                                            "WHERE Local <> 'NUM' " & _
                                            "ORDER BY Local ASC;")

                    if not (tsim.bof or tsim.eof) then
                        response.write "'*'"

                        do
                            response.write ", "
                            response.write "'" & tsim("simbolo") & "'"
                            tsim.MoveNext
                        loop until (tsim.eof)
                    end if

                    tsim.close: set tsim = nothing
                %>];

                var formula = [<%
                    set tfor = con.execute("SELECT Local, Simbolo, Formula " & _
                                            "FROM seg_Cripto_NumParse_Locales " & _
                                            "WHERE Local <> 'NUM' " & _
                                            "ORDER BY Local ASC;")

                    if not (tfor.bof or tfor.eof) then
                        response.write "'*'"

                        do
                            response.write ", "
                            response.write "'" & tfor("formula") & "'"
                            tfor.MoveNext
                        loop until (tfor.eof)
                    end if

                    tfor.close: set tfor = nothing
                %>];
            
                if (desde == hasta) {
                    if (m == 1) {document.getElementById("txtMontoCambio").value = monto};
                    if (m == 2) {document.getElementById("Monto").value = monto};
                } else {
                    /* Llevamos el monto a USD */

                    donde = 0;

                    for(let k = 0; k < locales.length; k++) {
                        if (locales[k] == desde) {
                            donde = k
                        }
                    };

                    monto2 = (monto / formula[donde]);

                    /* Llevamos el monto USD a Moneda Destino */

                    donde = 0;

                    for(let k = 0; k < locales.length; k++) {
                        if (locales[k] == hasta) {
                            donde = k
                        }
                    };

                    monto2 = (monto2 * formula[donde]);
                    monto2 = formatter.format(monto2);

                    if (m == 1) {document.getElementById("txtMontoCambio").value = monto2};
                    if (m == 2) {document.getElementById("Monto").value = monto2};
                }
            };

            function sumarUnDia(fechaStr) {
                /*
                    IMPORTANTE:

                    Esta funcion, tal como aparece a continuación, 
                    sólo funciona con el formato de fecha "dd/MM/aaaa"

                    Si el formato de fecha en la aplicación cambia,
                    la función fallará en su respuesta.
                */

                let partes = fechaStr.split("/");
                let dia = parseInt(partes[0], 10);
                let mes = parseInt(partes[1], 10) - 1; // Restamos 1 porque los meses en JavaScript van de 0 a 11
                let anio = parseInt(partes[2], 10);

                // Crear objeto Date y sumar un día
                let fecha = new Date(anio, mes, dia);
                fecha.setDate(fecha.getDate() + 1);

                // Formatear la nueva fecha en "dd/MM/aaaa"
                let nuevoDia = fecha.getDate().toString().padStart(2, "0");
                let nuevoMes = (fecha.getMonth() + 1).toString().padStart(2, "0");
                let nuevoAnio = fecha.getFullYear();

                return `${nuevoDia}/${nuevoMes}/${nuevoAnio}`;
            }

            function restarUnDia(fechaStr) {
                /*
                    IMPORTANTE:

                    Esta funcion, tal como aparece a continuación, 
                    sólo funciona con el formato de fecha "dd/MM/aaaa"
                    
                    Si el formato de fecha en la aplicación cambia,
                    la función fallará en su respuesta.
                */

                let partes = fechaStr.split("/");
                let dia = parseInt(partes[0], 10);
                let mes = parseInt(partes[1], 10) - 1; // Restamos 1 porque los meses en JavaScript van de 0 a 11
                let anio = parseInt(partes[2], 10);

                // Crear objeto Date y restar un día
                let fecha = new Date(anio, mes, dia);
                fecha.setDate(fecha.getDate() - 1);

                // Formatear la nueva fecha en "dd/MM/aaaa"
                let nuevoDia = fecha.getDate().toString().padStart(2, "0");
                let nuevoMes = (fecha.getMonth() + 1).toString().padStart(2, "0");
                let nuevoAnio = fecha.getFullYear();

                return `${nuevoDia}/${nuevoMes}/${nuevoAnio}`;
            }         
            
            function sumar30Minutos(hora) {
                /*
                    IMPORTANTE:

                    Esta funcion, tal como aparece a continuación, 
                    sólo funciona con el formato de hora es "HH:mm"
                    (o sea, formato de 24 horas con "0" como padding)
                    
                    Si el formato de hora en la aplicación cambia,
                    la función fallará en su respuesta.

                    La hora siempre se redondea a los 30 minutos
                */

                let [hh, mm] = hora.split(":").map(Number);
                let fecha = new Date();

                if (mm <= 15) { 
                    mm = 0; 
                } else {
                    mm = 30;
                };

                fecha.setHours(hh, mm);
                fecha.setMinutes(fecha.getMinutes() + 30);
                
                return fecha.toTimeString().slice(0, 5);
            }

            function restar30Minutos(hora) {
                /*
                    IMPORTANTE:

                    Esta funcion, tal como aparece a continuación, 
                    sólo funciona con el formato de hora es "HH:mm"
                    (o sea, formato de 24 horas con "0" como padding)
                    
                    Si el formato de hora en la aplicación cambia,
                    la función fallará en su respuesta.

                    La hora siempre se redondea a los 30 minutos
                */
                               
                let [hh, mm] = hora.split(":").map(Number);
                let fecha = new Date();
                
                if (mm <= 15) { 
                    mm = 0; 
                } else {
                    mm = 30;
                };

                fecha.setHours(hh, mm);
                fecha.setMinutes(fecha.getMinutes() - 30);
                
                return fecha.toTimeString().slice(0, 5);
            }                

            function swalert(mensaje) {
                Swal.fire({
                    text: mensaje,
                    icon: 'info',
                    background: '#ffffff',
                    color: '#222',
                    confirmButtonText: 'OK',
                    confirmButtonColor: '#3b82f6',  // azul moderno
                    backdrop: 'rgba(0,0,0,0.35)',
                    customClass: {
                        popup: 'swal-white-premium'
                    }
                });
            }

            mask(document.getElementById('txt_fecha'), ['99/99/9999']);
            mask(document.getElementById('txt_hora'),  ['99:99']);            
        </script>    

        <% con.close: set con = nothing %>
        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>