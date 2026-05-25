<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>

        <title>Editar Notas</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0090"
            SysLockOut

            '
            ' Funciones y Procedimientos
            '
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
                dim char, k, res

                res = ""

                for k = 1 to (len(trim(cadena)))
                    char = mid(cadena, k, 1)

                    select case asc(char)
                        case 225: char = "a"
                        case 193: char = "A"
                        case 233: char = "e"
                        case 232: char = "e"
                        case 201: char = "E"
                        case 237: char = "i"
                        case 205: char = "I"
                        case 243: char = "o"
                        case 211: char = "O"
                        case 250: char = "u"
                        case 218: char = "U"
                        case 209: char = "N"
                        case 241: char = "n"
                    end select

                    res = res  & char
                next

                limpiar = res
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
        %>      
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    
        <%
            dim con, t, p, sqlString, llave, usu, pre, vinculo
            dim Nota, multi, dia, ver, tipo, estatus, ordenado

            usu = Request.Cookies("usuario")
            pre = Request.QueryString("p")
            llave = Request.QueryString("l")
            multi = MultiPrecio(pre)

            dia = Request.QueryString("d")
            ver = Request.QueryString("v")
            tipo = Request.QueryString("t")
            estatus = Request.QueryString("e")
            ordenado = Request.QueryString("o")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            sqlString = "SELECT d.Llave, d.Fecha, d.Hora, d.Descripcion, d.Nota, d.NotaPre, d.NotaDonde " & _
                            "FROM dbo.pre_Presupuesto_Detalles AS d " & _
                        "WHERE Llave = " & llave & ";"

            set t = con.execute(sqlString)
                
                HoraTemp = RIGHT("0000" & t("Hora"), 4)
                Fecha = FechaForm(t("Fecha"))
                Hora = HoraForm(HoraTemp)
                Descripcion = t("Descripcion")
                Nota = t("Nota")
        %>  

        <br />

        <form name="formulario" id="formulario" method="post" action="pre_det_grabar_nota.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Editar Nota (<%= limpiar(NombrePresupuesto(usu, pre)) %>)
                </div>
                
                <div style="flex: 0 0 50%; text-align: right;">
                    <button type="button" class="form-btn verde normal" onclick="submit()">Grabar</button>    
                    <button type="button" class="form-btn rojo normal" onclick="volver()">Cancelar</button>     
                </div>
            </div>        

            <div class="main main-scroll">
                <div class="no-ver">
                    <input id="ordenamiento"  name="ordenamiento"     type="text" value="<%= oParan %>">
                    <input id="txt_fecha"     name="txt_fecha"        type="text" value="<%= fecha %>" />
                    <input id="txt_hora"      name="txt_hora"         type="text" value="<%= hora %>"  />

                    <input id="dia"           name="dia"              type="text" value="<%= dia %>"  />
                    <input id="ver"           name="ver"              type="text" value="<%= ver %>"  />
                    <input id="tipo"          name="tipo"             type="text" value="<%= tipo %>"  />
                    <input id="estatus"       name="estatus"          type="text" value="<%= estatus %>"  />
                    <input id="ordenado"      name="ordenado"         type="text" value="<%= ordenado %>"  />
                </div>

                <div class="line">
                    <label class="label normal">Fecha</label>
                    <input class="field normal" type="text" disabled value="<%= fecha & " " & hora %>">
                </div>

                <div class="line">
                    <div class="full section" style="background-color: rgb(245, 245, 245);">
                        <% if t("NotaPre") = 1 then %> 
                            <textarea class="field full" name="txtNota" id="txtNota" rows=10 cols=80
                                      style="font-family: courier; font-size: 16px; width:100%;"><%= t("Nota") %></textarea>  

                        <% else %>
                            <script src="/core/lib/tinymce/tinymce.min.js"></script>                        
                            
                            <textarea class="field full editor" id="txtNota" name="txtNota"> 
                                <%= t("Nota") %>
                            </textarea>

                            <script>
                                tinymce.init({
                                    entity_encoding : "raw",
                                    selector: '.editor',
                                    license_key: 'gpl',
                                    height: 450,
                                    branding: false,
                                    promotion: false,                                    
                                    language: 'es',
                                    language_url: '/core/includes/es.js', 
                                    plugins: 'anchor autolink charmap codesample emoticons image link lists media searchreplace table visualblocks wordcount ',
                                    toolbar: 'undo redo | blocks fontfamily fontsize | bold italic underline strikethrough | link image media table mergetags | addcomment showcomments | spellcheckdialog a11ycheck typography | align lineheight | checklist numlist bullist indent outdent | emoticons charmap | removeformat',
                                    mergetags_list: [
                                        { value: 'First.Name', title: '' },
                                        { value: 'Email', title: '' },
                                    ]
                                });
                            </script>                
                        <% end if %>
                    </div>
                </div>

                <div class="line">
                    <label class="label normal">Tipo</label>
                    <select class="field large" name="NotaPre" id="NotaPre">
                        <option value="0" <% if t("NotaPre") = "0" then response.write " selected" %>>Ver Nota en formato Extendido</option>
                        <option value="1" <% if t("NotaPre") = "1" then response.write " selected" %>>Forzar el formato de bloque en impresión</option>     
                    </select>  
                </div>

                <div class="line">
                    <label class="label normal">Alineación</label>
                    <select class="field large"  name="NotaDonde" id="NotaDonde">
                        <option value="I" <% if t("NotaDonde") = "I" then response.write " selected" %>>Imprimir Nota en Panel Izquierdo</option>
                        <option value="D" <% if t("NotaDonde") = "D" then response.write " selected" %>>Imprimir Nota en Panel Derecho</option>     
                    </select> 
                </div>                
            </div>

            <div class="no-ver">
                <input id="Llave" name="Llave" type="text" value="<%= Llave %>"         />
                <input id="txtMulti" name="txtMulti" type="text" value="<%= multi  %>"  />
                <input id="txtPre" name="txtPre" type="text" value="<%= pre %>"         />
            </div>                    
        </form>

        <script>
            function submit() {
                document.getElementById("formulario").submit();
            }

            function volver() {
                var vinculo = "<%
                    vinculo = "pre_det_editar"

                    if MonedaOrigen <> MonedaDestino then
                        vinculo = vinculo & "_m"
                    end if             

                    vinculo = vinculo & ".asp?p=" & pre & "&d=" & request.QueryString("d") & "&v=" & request.QueryString("v") & "&t=" & request.QueryString("t") & "&e=" & request.QueryString("e") & "&o=" & request.QueryString("o")

                    response.write vinculo                         
                %>";

                window.location.href = vinculo;
            }                       
        </script>

        <%
            t.close: set t = nothing
            con.close: set con = nothing    
        %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->    
    </body>
</html>