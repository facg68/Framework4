<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Editar Registro de Presupuesto</title>

        <!-- #include virtual = "/core/includes/kernel/head.inc" -->        

        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0090"
            SysLockOut
 
            function EstatusPresupuesto(Presupuesto)
                dim c, ta

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set ta = c.Execute("SELECT Estatus from pre_Presupuesto_Encabezado where (Usuario = '" & Request.Cookies("Usuario") & "') AND (Presupuesto = '" & presupuesto & "');")
                        if not (ta.bof or ta.eof) then
                            EstatusPresupuesto = ta("Estatus")
                        else
                            EstatusPresupuesto = ""
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
            end Function

            function NombrePresupuesto(Usuario, Presupuesto)
                dim c, ta

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set ta = c.Execute("SELECT nombre from pre_Presupuesto_Encabezado where (Usuario = '" & usuario & "') AND (Presupuesto = '" & presupuesto & "');")
                        if not (ta.bof or ta.eof) then
                            NombrePresupuesto = ta("nombre")
                        else
                            NombrePresupuesto = ""
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
            end Function

            Function FechaForm(FechaDB)
                dim a, m, d

                FechaForm = ""

                if not isnull(FechaDB) then
                    d = RIGHT("00" & day(FechaDB) ,2)
                    m = RIGHT("00" & month(FechaDB), 2)
                    A = year(FechaDB)

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

            Function MonedaUsuario(Usuario)
                dim fcon, f, sqlString

                sqlString = "SELECT usuLocal FROM seg_Usuarios WHERE usuCodigo = '" & Usuario & "';"

                set fcon = Server.CreateObject("ADODB.Connection")
                fcon.open Application("Conn")
                    set f = fcon.execute(sqlString)
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
                fcon.close: set fcon = nothing
            end Function

            function limpiar(cadena)
                limpiar = cadena
                
                limpiar = Replace(limpiar, "&#11013;", "<")
                limpiar = Replace(limpiar, "&#11157;", ">")
            end function   
        
            function preLocalOrigen(presupuesto)
                dim fcon, f

                set fcon = Server.CreateObject("ADODB.Connection")
                fcon.Open Application("Conn")
                    set f = fcon.execute("SELECT MonedaOrigen FROM pre_Presupuesto_Encabezado WHERE Presupuesto = '" & presupuesto & "' AND Usuario = '" & Request.Cookies("Usuario") & "'")
                        preLocalOrigen = f("MonedaOrigen")
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

            Function HoraVbs()
                dim h, m
                
                h = RIGHT("00" & Hour(Time()), 2)
                m = RIGHT("00" & Minute(Time()), 2)

                HoraVbs = h & ":" & m
            end function  

            Function MultiPrecio(Presupuesto)
                dim c, ta

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set ta = c.Execute("SELECT multiprecio from pre_Presupuesto_Encabezado where (Usuario = '" & Request.Cookies("Usuario")  & "') AND (Presupuesto = '" & presupuesto & "');")
                        if not (ta.bof or ta.eof) then
                            MultiPrecio = ta("multiprecio")
                        else
                            MultiPrecio = 0
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
            end function    

            function sysDateTimeOffset()
                dim c, ta

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set ta = c.Execute("SELECT dbo.sysDateTimeOffset() AS FechaHoraSistema;")
                        sysDateTimeOffset = ta("FechaHoraSistema")
                    ta.close: set ta = nothing
                c.close: set c = nothing      
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

            function Presupuesto_Llave(Llave)
                dim c, ta, sqlString

                sqlString = "SELECT Presupuesto " & _
                            "FROM dbo.pre_Presupuesto_Detalles " & _
                            "WHERE (Llave = " & Llave & ");"

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set ta = c.Execute(sqlString)
                        if not (ta.bof or ta.eof) then
                            Presupuesto_Llave = ta("Presupuesto")
                        else
                            Presupuesto_Llave = NULL
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
            end Function     
            
            function tipoPresupuesto(Presupuesto)
                dim c, ta, sqlString

                sqlString = "SELECT Tipo " & _
                            "FROM dbo.pre_Presupuesto_Encabezado " & _
                            "WHERE (Presupuesto = '" & Presupuesto & "');"

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set ta = c.Execute(sqlString)
                        if not (ta.bof or ta.eof) then
                            tipoPresupuesto = ta("Tipo")
                        else
                            tipoPresupuesto = NULL
                        end if
                    ta.close: set ta = nothing
                c.close: set c = nothing
            end Function              
        %>           
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    

        <%
            dim con, t, p, sqlString, cbox, llave, nuevo, usu, pre, llaveCal, dia

            dim Fecha, Hora, CuentaOrigen, CuentaDestino, MontoOrigen
            dim MontoDestino, MonedaOrigen, MonedaDestino, Descripcion
            dim Contacto, Aplicado, HoraTemp, Estatus, TipoPre, Archivado
            dim redirectVer, redirectTipo, redirectEstatus, redirectOrdenamiento

            llaveCal = Request.QueryString("llaveCal")

            if llaveCal = "" then
                usu = Request.Cookies("usuario")
                pre = Request.QueryString("p")
                dia = Request.QueryString("d")
                llave = Request.QueryString("l")
                Estatus = EstatusPresupuesto(pre)
                TipoPre = tipoPresupuesto(pre)

                redirectVer = Request.QueryString("v")
                redirectTipo = Request.QueryString("t")
                redirectEstatus = Request.QueryString("e")
                redirectOrdenamiento = Request.QueryString("o")
            else
                usu = Request.Cookies("usuario")
                pre = Presupuesto_Llave(llaveCal)
                dia = "*"
                llave = llaveCal
                Estatus = EstatusPresupuesto(pre)
                TipoPre = tipoPresupuesto(pre)

                redirectVer = "0"
                redirectTipo = "P"
                redirectEstatus = "1"
                redirectOrdenamiento = "PD"
            end if      

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            nuevo = 0

            if llave < 0  then 
                nuevo = 1

                Fecha = sysDateTimeOffset_FechaForm
                Hora = sysDateTimeOffset_HoraForm
                CuentaOrigen = "PRE-000"
                CuentaDestino = "SYS-000"
                MontoOrigen = 0.00
                MontoDestino = 0.00
                MontoCambio = 0.00
                MonedaOrigen = preLocalOrigen(pre)
                MonedaDestino = preLocalDestino(pre)
                Descripcion = "Nueva Transaccion"
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

        <br />

        <form name="form_transaccion" id="form_transaccion" method="post" action="pre_det_grabar.asp">
            <div class="no-ver">
                <input id="usuario"         name="usuario"          type="text" value="<%= usu %>"                   />
                <input id="presupuesto"     name="presupuesto"      type="text" value="<%= pre %>"                   /> 
                <input id="d"               name="d"                type="text" value="<%= dia %>"                   /> 
                <input id="t"               name="t"                type="text" value="<%= redirectTipo %>"          /> 
                <input id="v"               name="v"                type="text" value="<%= redirectVer %>"           /> 
                <input id="e"               name="e"                type="text" value="<%= redirectEstatus %>"       /> 
                <input id="o"               name="o"                type="text" value="<%= redirectOrdenamiento %>"  /> 
                <input id="MonedaOrigen"    name="MonedaOrigen"     type="text" value="<%= MonedaOrigen %>"          />
                <input id="MonedaDestino"   name="MonedaDestino"    type="text" value="<%= MonedaDestino %>"         />
                <input id="Nuevo"           name="Nuevo"            type="text" value="<%= Nuevo %>"                 />
                <input id="Llave"           name="Llave"            type="text" value="<%
                                                                                            if Nuevo = 0 then 
                                                                                                response.write Llave 
                                                                                            else
                                                                                                response.write "0"
                                                                                            end if 
                                                                                        %>"                           />
            </div>        

            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 70%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    <input style="background: transparent; width: 100%; font-size: 18px;" type="text" value="<%= "Presupuesto: " & limpiar(NombrePresupuesto(usu, pre)) %>" readonly />
                </div>
                
                <div style="flex: 0 0 30%; text-align: right;">
                    <%
                        if (Estatus > 0) AND (Archivado = 0) then
                            %><button class="form-btn verde" style="width: 100px; font-size: 16px; color: white;" type="button" onClick="Verificar()">Grabar</button>&nbsp;<%
                        end if

                        vinculo = "pre_det_editar.asp?p=" & pre & "&d=" & dia & "&v=" & redirectVer & "&t=" & redirectTipo & "&e=" & redirectEstatus & "&o=" & redirectOrdenamiento

                        response.write "<a href='" & vinculo & "'>"
                        response.write "<button type='button' class='form-btn rojo' style='width: 100px; font-size: 16px; color: white;'>Cancelar</button>"     
                        response.write "</a>"                               
                    %>                           
                </div>
            </div>                              

            <div class="main main-scroll">
                <div class="line">
                    <label class="label normal">Descripcion</label>
                    <input class="field xxl" style="background-color: rgba(221, 235, 245, 1);"
                            id="Descripcion" name="Descripcion" type="text" value="<%= limpiar(Descripcion) %>" />
                </div>

                <div class="line">
                    <label class="label normal">Cuenta Origen</label>
                    <label class="label xxl" style="padding: 0px;">
                        <select class="field xxl " name="CuentaOrigen" id="CuentaOrigen" style="background-color: rgba(251, 230, 230, 1);">
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

                        &nbsp;&nbsp;

                        <button class="form-btn gris tiny" style="height: 30px;" type="button" onClick="swapCuentas()"> ▼ </button>
                    </label>                        
                </div>      

                <div class="line">
                    <label class="label normal">Cuenta Destino</label>
                    <label class="label xxl" style="padding: 0px;">
                        <select class="field xxl" name="CuentaDestino" id="CuentaDestino">
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
                    <label class="label normal">Monto Origen</label>
                    <label class="label xxl" style="padding: 0px;">
                        <input class="field tiny" style="background-color: rgba(251, 230, 230, 1);" 
                                id="Monto" name="Monto" type="text" value="<%
                            if MontoOrigen >= 0 then
                                response.write formatNumber(MontoOrigen)
                            else
                                response.write FormatNumber( (-1 * MontoOrigen) )
                            end if
                        %>" OnChange="cMoneda1('<%= MonedaOrigen %>','<%= MonedaDestino %>', 1);" style="background-color; rgba(241, 212, 212, 1);"/>

                        &nbsp;&nbsp;

                        <select class="field normal" style="background-color: rgba(251, 230, 230, 1);"
                                name="MonedaOrigenDisplay" id="MonedaOrigenDisplay" disabled style="background-color; rgba(241, 212, 212, 1);" >
                            <%
                                sqlString = "SELECT [Local], Simbolo + '  (' + NombreListas + ')' AS NombreLocal " & _
                                                "FROM seg_Cripto_NumParse_Locales " & _
                                            "WHERE [Local] <> 'NUM' " & _
                                            "ORDER BY NombreLocal ASC;"

                                set cbox = con.execute(sqlString)

                                if not (cbox.bof or cbox.eof) then
                                    Do
                                        response.write "<option value='" & cbox("Local") & "' "
                                            if MonedaOrigen = cbox("Local") then 
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

                <div class="<%
                                if MonedaOrigen = MonedaDestino then
                                    response.write "no-ver"
                                else
                                    response.write "line"
                                end if
                            %>">
                    <label class="label normal">Monto Destino</label>
                    <label class="label large">
                        <input class="field tiny" id="txtMontoCambio" name="txtMontoCambio" type="text" value="<%
                            if MontoCambio => 0 then
                                response.write formatNumber(MontoCambio)
                            else
                                response.write FormatNumber( (-1 * MontoCambio) )
                            end if
                        %>" OnChange="cMoneda2('<%= MonedaDestino %>','<%= MonedaOrigen %>', 2);" />

                        &nbsp;&nbsp;

                        <select class="field normal" name="MonedaDestinoDisplay" id="MonedaDestinoDisplay" disabled>
                            <%
                                sqlString = "SELECT [Local], Simbolo + '  (' + NombreListas + ')' AS NombreLocal " & _
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
                    <label class="label normal">Fecha</label>
                    <label class="label normal" style="padding: 0px;">
                        <button class="form-btn gris tiny" style="height: 30px;" type="button" onClick="fechaMenos()"> ◀ </button>&nbsp;&nbsp;
                            <input class="field small" id="txt_fecha" name="txt_fecha" type="text" value="<%= fecha %>" placeholder="dd/mm/aaaa" style="width: 125px; font-size: 18px; text-align: center;" OnChange="FechaValida();" />&nbsp;&nbsp;
                        <button class="form-btn gris tiny" style="height: 30px;" type="button" onClick="fechaMas()"> ▶ </button>

                    </label>
                </div>

                <div class="line">
                    <label class="label normal">Hora</label>
                    <label class="label normal" style="padding: 0px;">
                        <button class="form-btn gris tiny" style="height: 30px;" type="button" onClick="horaMenos()">◀</button>&nbsp;&nbsp;
                            <input class="field tiny" id="txt_hora" name="txt_hora" type="text" value="<%= hora %>" placeholder="hh:mm" style="width: 70px; text-align: center;" OnChange="HoraValida();" />&nbsp;&nbsp;
                        <button class="form-btn gris tiny" style="height: 30px;" type="button" onClick="horaMas()">▶</button>                                            
                    </label>                        
                </div>

                <div class="line">
                    <label class="label normal">Contacto</label>
                    <select class="field large" style="background-color: rgba(221, 235, 245, 1);" name="Contacto" id="Contacto">
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
                    <label class="label normal">Aplicar</label>
                    <select class="field normal" name="Aplicado" id="Aplicado" >
                        <option value="0" <% if Aplicado = 0 then response.write " selected" %>>No Aplicado</option>
                        <option value="1" <% if Aplicado = 1 then response.write " selected" %>>Aplicado</option>
                    </select>                           
                </div>

                <div class="line">
                    <label class="label normal">Acción</label>
                    <select class="field xxl" name="preSiguiente" id="preSiguiente" class="foot">
                        <option value="*"><%
                            if TipoPre = "M" then
                                response.write "Transacción del Modelo Actual" 
                            else
                                response.write "Transacción del Presupuesto Actual" 
                            end if
                            %></option>
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
            </div>
        </form>

        <br /><br />

        <script>
            <%
                dim cc, tloc, tsim, tfor

                set cc = server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
            %>

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

            function Verificar() {
                if ((HoraValida() == 0) && (FechaValida() == 0) && (CantidadValida() == 0)) {
                    document.getElementById("form_transaccion").submit();
                }
                else {
                    alert("Algunos de los valores entrados no son correctos. Por favor, verifique y vuelva a intentarlo.");
                }
            };

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
                    alert("El valor de la Fecha no es válido. Por favor verifique.");
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
                }
                else {
                    alert("El valor de la Hora no es válido. Por favor verifique.");  
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

                if (c1 && c2) {
                    return 0;
                } else {
                    return 1;
                }        
            };

            function cMoneda1(desde, hasta, m) {
                var c1;
                var mOrigen = document.getElementById("Monto").value;

                mOrigen = mOrigen.replace(/,/g, "");
                c1 = !isNaN(mOrigen); 

                if (c1) {
                    CambiarMoneda(desde, hasta, m);
                } else {
                    alert("El valor del Monto es incorrecto. Verifique y vuelva a intentarlo.");

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
                alert("El valor del Monto Destino es incorrecto. Verifique y vuelva a intentarlo.");
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
                        alert("El Valor SIEMPRE debe ser mayor o igual a cero.");
                        document.getElementById("Monto").value = monto;
                    };          
                };

                if (m == 2) {
                    monto = document.getElementById("txtMontoCambio").value

                    if (monto < 0) {
                        monto = (-1 * monto);
                        alert("El Valor SIEMPRE debe ser mayor o igual a cero.");
                        document.getElementById("txtMontoCambio").value = monto;
                    };          
                };

                var formatter = new Intl.NumberFormat('en-US', {
                    style: 'decimal',
                    currency: 'USD',
                });   

                var locales = [<%
                    set tloc = cc.execute("SELECT Local, Simbolo, Formula " & _
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
                    set tsim = cc.execute("SELECT Local, Simbolo, Formula " & _
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
                    set tfor = cc.execute("SELECT Local, Simbolo, Formula " & _
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
                } 
                else {
                    /*
                        Llevamos el monto a USD
                    */

                    donde = 0;

                    for(let k = 0; k < locales.length; k++) {
                        if (locales[k] == desde) {
                        donde = k
                        }
                    };

                    monto2 = (monto / formula[donde]);

                    /*
                        Llevamos el monto USD a Moneda Destino
                    */

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

            mask(document.getElementById('txt_fecha'), ['99/99/9999']);
            mask(document.getElementById('txt_hora'),  ['99:99']);
        </script>

        <% cc.close: set cc = nothing  %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->           
    </body>
</html>