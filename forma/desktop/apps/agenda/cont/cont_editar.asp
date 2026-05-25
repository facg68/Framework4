<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->   
        <title>Editar Contactos</title>
        
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0050"
            SysLockOut


            ' -- Init() --
                dim con, t, sqlString, cbox, c, nomContacto, nuevo
                dim ver, tipo, categ, orden1, orden2, tabActual

                dim usuario, codigo, primerNombre, segundoNombre, primerApellido, segundoApellido, correoElectronico, fechaCumple, estatus
                dim empresa, telefonoEmpresa, sitioWeb, pais, provincia, ciudad, direccion, notas, deSistema, visible, arbol, signo
                
                usuario = Request.Cookies("usuario")
                codigo = Request.QueryString("con")
                ver = Request.QueryString("v")
                tipo = Request.QueryString("t")
                categ = Request.QueryString("c")
                orden1 = Request.QueryString("o1")
                orden2 = Request.QueryString("o2")

                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")

                if codigo = ""  then 
                    codigo = CodigoNuevoContacto(usuario)
                    DeSistema = 0
                    Visible = 1
                    Arbol = 99
                    Signo = 99

                    nomContacto = ""
                    nuevo = 1      
                else
                    sqlString = "SELECT codigo, primerNombre, segundoNombre, primerApellido, segundoApellido, " & _
                                    " correoElectronico, fechaCumple, empresa, telefonoEmpresa, sitioWeb, pais, provincia, " & _
                                    " ciudad, direccion, notas, deSistema, visible, arbol, signo, Estatus " & _
                                "FROM con_Contactos as c " & _
                                "WHERE (Usuario = '" & usuario & "') " & _
                                "AND (Codigo = '" & codigo & "');"

                    set t = con.execute(sqlString)
                        primerNombre = t("primerNombre")
                        segundoNombre = t("segundoNombre")
                        primerApellido = t("primerApellido")        
                        segundoApellido = t("segundoApellido")
                        correoElectronico = t("correoElectronico")
                        fechaCumple = t("fechaCumple")
                        empresa = t("empresa")
                        telefonoEmpresa = t("telefonoEmpresa")
                        sitioWeb = t("sitioWeb")
                        pais = t("pais")
                        provincia = t("provincia")
                        ciudad = t("ciudad")
                        direccion = t("direccion")
                        notas = t("notas")
                        deSistema = t("deSistema")
                        visible = t("visible")
                        arbol = t("arbol")
                        signo = t("signo")
                        estatus = t("estatus")

                        nomContacto = ""
                        if PrimerNombre    <> "" then append nomContacto, " " & PrimerNombre        
                        if SegundoNombre   <> "" then append nomContacto, " " & SegundoNombre
                        if PrimerApellido  <> "" then append nomContacto, " " & PrimerApellido
                        if SegundoApellido <> "" then append nomContacto, " " & SegundoApellido

                        nuevo = 0
                    t.close: set t = nothing
                end if   
            ' -- Fin: Init() --

            ' -- Funciones y Procedimientos --
                function SignoUsuario()
                    dim zCon, zTab, sqlString

                    sqlString = "SELECT usuSigno " & _
                                "FROM seg_Usuarios " & _
                                "WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';"

                    set zCon = Server.CreateObject("ADODB.Connection")
                    zCon.open Application("Conn")
                        set zTab = zCon.execute(sqlString)
                            SignoUsuario = ztab("usuSigno")
                        ztab.close: set ztab= nothing
                    zCon.close: set zCon = nothing
                end function     

                sub append(byRef Cadena, NuevaCadena)
                    if NuevaCadena <> "" then
                        Cadena = Cadena & NuevaCadena
                    end if
                end sub                   

                function SecObjeto(Usuario, Contacto)
                    dim aCon, aTab, sqlString

                    sqlString = "SELECT dbo.con_Adjuntos_Secuencia('" & Usuario & "', '" & Contacto & "') AS AdjuntoSec;"

                    set aCon = Server.CreateObject("ADODB.Connection")
                    aCon.open Application("Conn")
                        set aTab = aCon.execute(sqlString)
                            SecObjeto = atab("AdjuntoSec")
                        atab.close: set atab = nothing
                    aCon.close: set aCon = nothing          
                end function

                function Categs(Tipo, Contacto)
                    dim cc, tt, sqlString

                    sqlString = "SELECT c.Nombre " & _
                                "FROM dbo.con_Contactos_Categorias AS c " & _
                            "INNER JOIN dbo.con_Contactos_ConCategs AS cc " & _
                                    "ON c.Codigo = cc.Categoria " & _
                                "WHERE (c.Usuario = '" & Request.Cookies("Usuario") &"') " & _
                                "AND c.Tipo = '" & Tipo & "' " & _
                                "AND (cc.Codigo = '" & Contacto & "') " & _
                            "ORDER BY c.Nombre;"

                    set cc = Server.CreateObject("ADODB.Connection")
                    cc.open Application("Conn")        
                        set tt = cc.execute(sqlString)
                            Categs = ""
                            
                            if not (tt.bof or tt.eof) then
                                Do
                                    Categs = Categs & tt("Nombre") & ", " 

                                    tt.MoveNext
                                Loop Until tt.eof

                                Categs = left(Categs, (Len(Categs) -2))
                            end if
                        tt.close: set tt = nothing
                    cc.close: set cc = nothing
                end function  

                function CodigoNuevoContacto(Usuario)
                    dim cc, tt, sqlString, sec

                    sqlString = "SELECT MAX(CAST(Codigo AS Numeric(12, 0))) AS M " & _
                                "FROM dbo.con_Contactos " & _
                                "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                "AND (Codigo NOT IN (SELECT usuCodigo FROM dbo.seg_Usuarios));"

                    set cc = Server.CreateObject("ADODB.Connection")
                    cc.open Application("Conn")        
                        set tt = cc.execute(sqlString)
                            if (tt.bof or tt.eof) then
                                CodigoNuevoContacto = "00000000001"
                            else
                                sec = cInt(tt("m")) + 1
                                CodigoNuevoContacto = right("000000000000" & sec, 12)
                            end if
                        tt.close: set tt = nothing
                    cc.close: set cc = nothing
                end function     

                function FechaFormulario(FechaServer)  
                    dim d, m, a, f

                    if FechaServer <> "" then
                        f = cstr(FechaServer)

                        a = Left(f, 4)
                        m = right("00" & Mid(f, 6, 2), 2)
                        d = right("00" & Right(f, 2), 2)

                        FechaFormulario = d & "/" & m & "/" & a
                    end if
                end function

                function HoraFormulario(HoraServer)  
                    dim hh, h, m

                    if HoraServer <> "" then
                        hh = right("0000" & cStr(HoraServer), 4)

                        h = left(hh, 2)
                        m = right(hh, 2)

                        HoraFormulario = h & ":" & m
                    else
                        HoraFormulario = "00:00"
                    end if
                end function    

                Function TipoContacto(Codigo)
                    dim cc, tt, sqlString, sec

                    sqlString = "SELECT TOP (1) cc.Tipo " & _
                                "FROM con_Contactos AS c INNER JOIN con_Contactos_ConCategs AS cc " & _
                                "ON c.Usuario = cc.Usuario AND c.Codigo = cc.Codigo " & _
                                "WHERE(c.Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                "AND (c.Codigo = '" & Codigo & "')"

                    set cc = Server.CreateObject("ADODB.Connection")
                    cc.open Application("Conn")        
                        set tt = cc.execute(sqlString)
                            if (tt.bof or tt.eof) then
                                TipoContacto = "*"
                            else
                                TipoContacto = tt("Tipo")
                            end if
                        tt.close: set tt = nothing
                    cc.close: set cc = nothing
                End Function                
            ' -- Fin: Funciones y Procedimientos

            ' -- Secciones --
                sub tab_Generales()
                    %>
                        <div class="no-ver">
                            <input id="cod" name="cod" type="text" value="<%= codigo %>" />
                            <input id="usu" name="usu" type="text" value="<%= usuario %>" />
                            <input id="param_ver"     name="param_ver"      type="text" value="<%= ver %>" />
                            <input id="param_tipo"    name="param_tipo"     type="text" value="<%= tipo %>" />
                            <input id="param_categ"   name="param_categ"    type="text" value="<%= categ %>" />
                            <input id="param_orden1"  name="param_orden1"   type="text" value="<%= orden1 %>" />
                            <input id="param_orden2"  name="param_orden2"   type="text" value="<%= orden2 %>" />                        
                        </div>

                        <div class="line label-top">
                            <label class="label tiny2" style="vertical-align: top;">Generales</label>
                            <div class="label full section">
                                <div class="line">
                                    <label class="label normal">Primer Nombre</label>
                                    <input class="field normal" id="primerNombre" name="primerNombre" type="text" value="<%= primerNombre %>" />
                                </div>

                                <div class="line">
                                    <label class="label normal">Segundo Nombre</label>
                                    <input class="field normal" id="segundoNombre" name="segundoNombre" type="text" value="<%= segundoNombre %>" />
                                </div>

                                <div class="line">
                                    <label class="label normal">Primer Apellido</label>
                                    <input class="field normal" id="primerApellido" name="primerApellido" type="text" value="<%= primerApellido %>" />
                                </div>

                                <div class="line">
                                    <label class="label normal">Segundo Apellido</label>
                                    <input class="field normal" id="segundoApellido" name="segundoApellido" type="text" value="<%= segundoApellido %>" />
                                </div>   

                                <div class="line">
                                    <label class="label normal">Correo</label>
                                    <input class="field normal" id="correoElectronico" name="correoElectronico" type="text" value="<%= correoElectronico %>" />
                                </div>   

                                <div class="line">
                                    <label class="label normal">Cumpleaños</label>
                                    <input class="field tiny" id="fechaCumple" name="fechaCumple" type="text" value="<%= fechaCumple %>" placeholder="dd/mm" />
                                </div>

                                <div class="line">
                                    <label class="label normal">Empresa</label>
                                    <input class="field normal" id="empresa" name="empresa" type="text" value="<%= empresa %>" />
                                </div>

                                <div class="line">
                                    <label class="label normal">Teléfono Empresa</label>
                                    <input class="field small" id="telefonoEmpresa" name="telefonoEmpresa" type="text" value="<%= telefonoEmpresa %>" />
                                </div>

                                <div class="line">
                                    <label class="label normal">Sitio Web</label>
                                    <input class="field normal" id="sitioWeb" name="sitioWeb" type="text" value="<%= sitioWeb %>" />
                                </div>
                            </div>
                        </div>
                    <%     
                end sub   
                
                sub tab_Telefonos()
                    %>
                        <div class="line label-top">
                            <label class="label tiny2" style="vertical-align: top;">Teléfonos</label>
                            <div class="label full section">
                                <div class="tabla-wrapper">
                                    <table class="tabla tabla-verde">
                                        <thead>
                                            <tr>
                                                <th class="sticky" style="width: 40%; text-align: center;">Número</th>
                                                <th class="sticky" style="width: 50%; text-align: center;">Tipo</th>
                                                <th class="sticky" style="width: 10%; text-align: center;">&nbsp;</th>
                                            </tr>
                                        </thead>  

                                        <tbody>                              
                                            <%
                                                sqlString = "SELECT t.Secuencia, t.Telefono, c.Nombre AS NombreTipo, c.Tipo " & _
                                                              "FROM con_Contactos_Telefonos AS t " & _
                                                        "INNER JOIN con_Contactos_Telefonos_Tipos AS c " & _
                                                                "ON t.Tipo = c.Tipo " & _
                                                             "WHERE (t.Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                                               "AND (t.Codigo = '" & Codigo & "') " & _
                                                          "ORDER BY c.Tipo, t.Secuencia;"

                                                set cbox = con.execute(sqlString)
                                                    if not (cbox.bof or cbox.eof) then
                                                        Cuantos = 0

                                                        Do
                                                            Cuantos = Cuantos + 1

                                                            %>
                                                                <tr>
                                                                    <td> <%= cbox("Telefono") %> </td>
                                                                    <td> <%= cbox("NombreTipo") %> </td>
                                                                    <td style="width: 10%; text-align: center;"> 
                                                                        <button type="button" 
                                                                                class="form-btn rojo" 
                                                                                onclick="BorrarTel('<%= cbox("Secuencia") %>','<%= cbox("Telefono") %>')" >
                                                                            <i class="fa fa-trash fa-xl"></i>
                                                                        </button>
                                                                    </td>
                                                                </tr>
                                                            <%

                                                            cbox.MoveNext
                                                        Loop Until cbox.eof
                                                    end if
                                                cbox.close: set cbox = nothing
                                            %>

                                            <!--
                                                Añadimos un "formulario" para añadir 
                                                más títulos...
                                            -->                 

                                            <tr>
                                                <td>
                                                    <input class="field frame" 
                                                           style="width: 100%;" 
                                                           id="nuevoTelefono" name="nuevoTelefono" 
                                                           type="text" placeholder="9999-9999" />
                                                </td>

                                                <td>
                                                    <select class="field frame" 
                                                            style="width: 100%;"
                                                            name='nuevoTipo' id='nuevoTipo' >
                                                        <%
                                                            sqlString = "SELECT Tipo, Nombre " & _
                                                                          "FROM con_Contactos_Telefonos_Tipos AS c " & _
                                                                         "WHERE ( Tipo NOT IN " & _
                                                                                        "(SELECT Tipo " & _
                                                                                           "FROM con_Contactos_Telefonos " & _
                                                                                          "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                                                                            "AND (Codigo = '" & Codigo & "'))" & _ 
                                                                                ") " & _
                                                                      "ORDER BY Tipo;"

                                                            set cbox = con.execute(sqlString)
                                                                if not (cbox.bof or cbox.eof) then
                                                                    Do
                                                                        response.write "<option value='" & cbox("Tipo") & "'>" & cbox("Nombre") & "</option>"
                                                                        cbox.MoveNext
                                                                    Loop Until cbox.eof
                                                                else
                                                                    response.write "<option value='5'>Otro</option>"
                                                                end if

                                                            cbox.close: set cbox = nothing
                                                        %>
                                                    </select>              
                                                </td>

                                                <td style="text-align: center;">
                                                    <button class="form-btn verde" 
                                                            type="button"
                                                            onclick="NuevoTel()" >
                                                        <i class="fa fa-save fa-xl"></i>
                                                    </button>
                                                </td>
                                            </tr>
                                        </tbody>

                                        <tfoot>
                                            <tr>
                                                <td colspan="3" class="sticky" style="text-align: center;">
                                                    <%
                                                        select case Cuantos
                                                            case 0: response.write "El Contacto No Tiene Teléfonos"
                                                            case 1: response.write "El Contacto Tiene Un Teléfono"
                                                            case else
                                                                response.write "El Contacto Tiene " & Cuantos & " Teléfonos"
                                                        end select
                                                    %>
                                                </td>
                                            </tr>
                                        </tfoot>
                                    </table> 
                                </div>
                            </div>
                        </div>
                    <%     
                end sub  

                sub tab_Direcciones()
                    %>
                        <div class="line label-top" id="tab_Direcciones">
                            <label class="label tiny2" style="vertical-align: top;">Direcciones</label>
                            <div class="label full section">
                                <div class="line-group">
                                    <div class="line">
                                        <label class="label small">País</label>
                                        <select class="field large" name="cboPais" id="cboPais" >
                                            <%
                                                sqlString ="SELECT Codigo, Nombre FROM seg_Paises ORDER BY Nombre ASC;"

                                                set cbox = con.execute(sqlString)

                                                if not (cbox.bof or cbox.eof) then
                                                    Do
                                                        response.write "<option value='" & cbox("Codigo") & "' "
                                                            if pais = cbox("Codigo") then 
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
                                        <label class="label small">Provincia</label>
                                        <input class="field large" id="Provincia" name="Provincia" type="text" value="<%= provincia %>" />
                                    </div>            

                                    <div class="line">
                                        <label class="label small">Ciudad</label>
                                        <input class="field large" id="ciudad" name="ciudad" type="text" value="<%= ciudad %>" />
                                    </div>  

                                    <div class="line">
                                        <label class="label small">Dirección</label>
                                        <textarea class="field" style="width: 90%;"  
                                                name="txtAreaDireccion" id="txtAreaDireccion" 
                                                rows=5 cols=60 class="vbControl_Enabled item" 
                                            ><%= direccion %></textarea>
                                    </div>
                                </div>
                            </div>
                        </div>
                    <%     
                end sub  

                sub tab_Categorias()
                    %>
                        <div class="line label-top" id="tab_Categorias">
                            <label class="label tiny2" style="vertical-align: top;">Tipos</label>
                            <div class="label full section">
                                <div class="tabla-wrapper">
                                    <table class="tabla tabla-verde">
                                        <thead>
                                            <tr>
                                                <th class="sticky" style="width: 90%; text-align: center;">Categoría</th>
                                                <th class="sticky" style="width: 10%; text-align: center;">&nbsp;</th>
                                            </tr>
                                        </thead>  

                                        <tbody>                              
                                            <%
                                                sqlString = "SELECT c.Usuario, c.Codigo, cc.Tipo, cc.Categoria, cat.Nombre " & _
                                                            "FROM dbo.con_Contactos AS c " & _
                                                            "INNER JOIN dbo.con_Contactos_ConCategs AS cc " & _
                                                            "ON c.Usuario = cc.Usuario AND c.Codigo = cc.Codigo " & _
                                                            "INNER JOIN dbo.con_Contactos_Categorias AS cat " & _
                                                            "ON cc.Usuario = cat.Usuario AND cc.Tipo = cat.Tipo AND cc.Categoria = cat.Codigo " & _
                                                            "WHERE (c.Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                                            "AND (c.Codigo = '" & Codigo & "') " & _
                                                            "ORDER BY cat.Nombre;"

                                                set cbox = con.execute(sqlString)
                                                    if not (cbox.bof or cbox.eof) then
                                                        Cuantos = 0

                                                        Do
                                                            Cuantos = Cuantos + 1
                                                            borrarVinculo = "'" & cbox("Codigo") & "', '" & cbox("Tipo") & "', '" & cbox("Categoria") & "', '" & cbox("Nombre") & "'"
                                                            %>
                                                                <tr>
                                                                    <td><%= cbox("Nombre") %></td>

                                                                    <td style="text-align: center;">
                                                                        <button class = "form-btn rojo" 
                                                                                type = "button"
                                                                                onclick="BorrarCat(<%= borrarVinculo %>)" >
                                                                            <i class="fa fa-trash fa-xl"></i>
                                                                        </button>
                                                                    </td>
                                                                </tr>
                                                            <%

                                                            cbox.MoveNext
                                                        Loop Until cbox.eof
                                                    end if
                                                cbox.close: set cbox = nothing
                                            %>

                                            <!--
                                                Añadimos un "formulario" para añadir 
                                                más títulos...
                                            -->                 

                                            <tr>
                                                <td>
                                                    <%
                                                        sqlString = "SELECT Codigo, Nombre " & _
                                                                    "FROM con_Contactos_Categorias AS c " & _
                                                                    "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                                                    "AND (c.Tipo = '" & TipoContacto(Codigo) & "') " & _
                                                                    "AND ( Codigo NOT IN " & _
                                                                                "(SELECT Categoria " & _
                                                                                 "FROM con_Contactos_ConCategs " & _
                                                                                 "WHERE Usuario = '" & Request.Cookies("Usuario") & "' " & _
                                                                                 "AND Tipo = '" & TipoContacto(Codigo) & "' " & _
                                                                                 "AND Codigo = '" & Codigo & "')" & _ 
                                                                        ") " & _
                                                                    "ORDER BY Nombre;"                                                                                                                  
                                                    %>

                                                    <select class="field frame" 
                                                            style="width: 100%;"
                                                            name='frm_NuevaCateg' 
                                                            id='frm_NuevaCateg'>
                                                        <option value='*'>&nbsp;</option>

                                                        <%
                                                            set cbox = con.execute(sqlString)                                     
                                                                Do
                                                                    response.write "<option value='" & cbox("Codigo") & "'>" & cbox("Nombre") & "</option>"
                                                                    cbox.MoveNext
                                                                Loop Until cbox.eof 
                                                            cbox.close: set cbox = nothing
                                                        %>
                                                    </select>
                                                </td>

                                                <td style="text-align: center;">
                                                    <button class="form-btn verde"
                                                            type="button" 
                                                            onclick="NuevaCat('<%= Tipo %>')">
                                                        <i class="fa fa-save fa-xl"></i>
                                                    </button>                                       
                                                </td>
                                            </tr>
                                        </tbody>

                                        <tfoot>
                                            <tr>
                                                <td colspan="2" class="sticky" style="text-align: center;">
                                                    <%
                                                        select case Cuantos
                                                            case 0: response.write "El Contacto No Tiene Categorías Asignadas"
                                                            case 1: response.write "El Contacto Tiene Una Categorías Asignada"
                                                            case else
                                                                response.write "El Contacto Tiene " & Cuantos & " Categorías Asignadas"
                                                        end select
                                                    %>
                                                </td>
                                            </tr>
                                        </tfoot>
                                    </table> 
                                </div>
                            </div>
                        </div>
                    <%     
                end sub

                sub tab_ContactosRelacionados()
                    %>
                        <div class="line label-top">
                            <label class="label tiny2" style="vertical-align: top;">Contactos Relacionados</label>
                            <div class="label full section">
                                <div class="tabla-wrapper">
                                    <table class="tabla tabla-verde">
                                        <thead>
                                            <tr>
                                                <th class="sticky" style="width:  20%; text-align: center;">Tipo</th>
                                                <th class="sticky" style="width:  60%; text-align: center;">Contactos</th>
                                                <th class="sticky" style="width:  10%; text-align: center;">Cumple.</th>
                                                <th class="sticky" style="width:  10%; text-align: center;">&nbsp;</th>
                                            </tr>
                                        </thead>  

                                        <tbody>                              
                                            <%
                                                sqlString = "SELECT cr.Secuencia, r.NomRelacion, dbo.con_Contactos_NombreContacto(cr.Usuario, cr.CodigoRelacionado) AS Nombre, cr.Cumple " & _
                                                            "FROM con_Contactos_Relacionados AS cr " & _
                                                        "INNER JOIN con_Contactos_Relaciones AS r " & _
                                                                "ON cr.Relacion = r.CodRelacion " & _
                                                            "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                                            "AND (Codigo = '" & Codigo & "') " & _
                                                        "ORDER BY Nombre ASC;"

                                                set cbox = con.execute(sqlString)
                                                    if not (cbox.bof or cbox.eof) then
                                                        Cuantos = 0

                                                        Do
                                                            Cuantos = Cuantos + 1

                                                            %>
                                                                <tr>
                                                                    <td> <%= cbox("NomRelacion") %> </td>
                                                                    <td> <%= cbox("Nombre") %> </td>
                                                                    <td style="text-align: center;"> <%= cbox("Cumple") %> </td>
                                                                    <td style="text-align: center;"> 
                                                                        <button type="button" 
                                                                                class="form-btn rojo" 
                                                                                onclick="BorrarContRelacionado('<%= cbox("Secuencia") %>','<%= cbox("Nombre") %>')" >
                                                                            <i class="fa fa-trash fa-xl"></i>
                                                                        </button>
                                                                    </td>
                                                                </tr>
                                                            <%

                                                            cbox.MoveNext
                                                        Loop Until cbox.eof
                                                    end if
                                                cbox.close: set cbox = nothing
                                            %>

                                            <!--
                                                Añadimos un "formulario" para añadir 
                                                más títulos...
                                            -->                 

                                            <tr>
                                                <td style="width:20%;">
                                                    <select class="field frame" 
                                                            style="width: 100%;"
                                                            name='NuevaRelacion' id='NuevaRelacion' >
                                                        <%
                                                            sqlString = "SELECT CodRelacion, NomRelacion " & _
                                                                        "FROM con_Contactos_Relaciones AS cr;"

                                                            set cbox = con.execute(sqlString)
                                                                if not (cbox.bof or cbox.eof) then
                                                                    Do
                                                                        response.write "<option value='" & cbox("CodRelacion") & "'>" & cbox("NomRelacion") & "</option>"
                                                                        cbox.MoveNext
                                                                    Loop Until cbox.eof
                                                                end if
                                                            cbox.close: set cbox = nothing
                                                        %>  
                                                    </select>              
                                                </td>

                                                <td>
                                                    <select class="field frame" 
                                                            style="width: 100%;"
                                                            name='NuevoContactoRelacionado' 
                                                            id='NuevoContactoRelacionado' >
                                                        <option value='*'>- - SELECCIONAR - -</option>
                                                        <%
                                                            sqlString = "SELECT Codigo, dbo.con_Contactos_NombreContacto(Usuario, Codigo) AS Nombre, FechaCumple " & _
                                                                        "FROM con_Contactos AS c " & _
                                                                        "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                                                        "AND (dbo.con_IncluidoEnTipo(Usuario, Codigo, '" & tipo & "') = 1) " & _
                                                                        "AND (Codigo NOT IN (SELECT CodigoRelacionado " & _
                                                                                                "FROM con_Contactos_Relacionados AS cr " & _
                                                                                                "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                                                                                "AND (Codigo = '" & Codigo & "'))) " & _
                                                                    "ORDER BY Nombre;"

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
                                                </td>

                                                <td>
                                                    <input class="field frame" 
                                                        style="width: 100%;"
                                                        id="NuevoCumpleContactoRelacionado" 
                                                        name="NuevoCumpleContactoRelacionado" 
                                                        type="text" 
                                                        placeholder="DD/MM" />
                                                </td>

                                                <td style="text-align: center;">
                                                    <button class="form-btn verde" 
                                                            type="button"
                                                            onclick="NuevoContRelacionado()">
                                                        <i class="fa fa-save fa-xl"></i>
                                                    </button>
                                                </td>
                                            </tr>
                                        </tbody>

                                        <tfoot>
                                            <tr>
                                                <td colspan="4" class="sticky" style="text-align: center;">
                                                    <%
                                                        select case Cuantos
                                                            case 0: response.write "El Contacto No Tiene Contactos Relacionados"
                                                            case 1: response.write "El Contacto Tiene Una Contacto Relacionado"
                                                            case else
                                                                response.write "El Contacto Tiene " & Cuantos & " Contactos Relacionados"
                                                        end select
                                                    %>
                                                </td>
                                            </tr>
                                        </tfoot>
                                    </table> 
                                </div>
                            </div>
                        </div>
                    <%     
                end sub      

                sub tab_ContactosNoRelacionados()
                    %>
                        <div class="line label-top">
                            <label class="label tiny2" style="vertical-align: top;">Contactos Externos</label>
                            <div class="label full section">
                                <div class="tabla-wrapper">
                                    <table class="tabla tabla-verde">
                                        <thead>
                                            <tr>
                                                <th class="sticky" style="width:  20%; text-align: center;">Tipo</th>
                                                <th class="sticky" style="width:  60%; text-align: center;">Contactos</th>
                                                <th class="sticky" style="width:  10%; text-align: center;">Cumple.</th>
                                                <th class="sticky" style="width:  10%; text-align: center;">&nbsp;</th>                                                
                                            </tr>
                                        </thead>  

                                        <tbody>                              
                                            <%
                                                sqlString = "SELECT nr.Secuencia, nr.Usuario, nr.Codigo, nr.Relacion, nr.Nombre, nr.Cumple, r.NomRelacion " & _
                                                            "FROM con_Contactos_No_Relacionados AS nr " & _
                                                            "INNER JOIN con_Contactos_Relaciones AS r " & _
                                                            "ON nr.Relacion = r.CodRelacion " & _
                                                            "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                                            "AND (Codigo = '" & Codigo & "') " & _
                                                            "ORDER BY Nombre ASC;"

                                                set cbox = con.execute(sqlString)
                                                    if not (cbox.bof or cbox.eof) then
                                                        Cuantos = 0

                                                        Do
                                                            Cuantos = Cuantos + 1

                                                            %>
                                                                <tr>                                                                
                                                                    <td> <%= cbox("NomRelacion") %> </td>
                                                                    <td> <%= cbox("Nombre") %> </td>
                                                                    <td style="text-align: center;"> <%= cbox("Cumple") %> </td>
                                                                    <td style="text-align: center;"> 
                                                                        <button type="button" 
                                                                                class="form-btn rojo" 
                                                                                onclick="BorrarContNoRelacionado('<%= cbox("Secuencia") %>','<%= cbox("Nombre") %>')" >
                                                                            <i class="fa fa-trash fa-xl"></i>
                                                                        </button>
                                                                    </td>
                                                                </tr>
                                                            <%

                                                            cbox.MoveNext
                                                        Loop Until cbox.eof
                                                    end if
                                                cbox.close: set cbox = nothing
                                            %>

                                            <!--
                                                Añadimos un "formulario" para añadir 
                                                más títulos...
                                            -->                 

                                            <tr>
                                                <td style="width:20%;">
                                                    <select class="field frame" 
                                                            style="width: 100%;"
                                                            name='NuevaNoRel' id='NuevaNoRel' >
                                                        <%
                                                            sqlString = "SELECT CodRelacion, NomRelacion " & _
                                                                        "FROM con_Contactos_Relaciones AS cr;"

                                                            set cbox = con.execute(sqlString)
                                                                if not (cbox.bof or cbox.eof) then
                                                                    Do
                                                                        response.write "<option value='" & cbox("CodRelacion") & "'>" & cbox("NomRelacion") & "</option>"
                                                                        cbox.MoveNext
                                                                    Loop Until cbox.eof
                                                                end if
                                                            cbox.close: set cbox = nothing
                                                        %>  
                                                    </select>              
                                                </td>

                                                <td>
                                                    <input class="field frame" 
                                                        style="width: 100%;"
                                                        id="NuevoContNoRel" name="NuevoContNoRel" 
                                                        type="text" value="" 
                                                        placeholder="Nuevo Nombre" />
                                                </td>

                                                <td>
                                                    <input class="field frame" 
                                                        style="width: 100%;"
                                                        id="NuevoCumpleContNoRel" 
                                                        name="NuevoCumpleContNoRel" 
                                                        type="text" 
                                                        placeholder="DD/MM" />
                                                </td>

                                                <td style="text-align: center;">
                                                    <button class="form-btn verde" 
                                                            type="button"
                                                            onclick="NuevoContNoRelacionado()" >
                                                        <i class="fa fa-save fa-xl"></i>
                                                    </button>
                                                </td>
                                            </tr>
                                        </tbody>

                                        <tfoot>
                                            <tr>
                                                <td colspan="4" class="sticky" style="text-align: center;">
                                                    <%
                                                        select case Cuantos
                                                            case 0: response.write "El Contacto No Tiene Contactos Externos"
                                                            case 1: response.write "El Contacto Tiene Una Contacto Externo"
                                                            case else
                                                                response.write "El Contacto Tiene " & Cuantos & " Contactos Externos"
                                                        end select
                                                    %>
                                                </td>
                                            </tr>
                                        </tfoot>
                                    </table> 
                                </div>
                            </div>
                        </div>
                    <%     
                end sub      

                sub tab_Adjuntos()
                    %>
                        <div class="line label-top">
                            <label class="label tiny2" style="vertical-align: top;">Adjuntos</label>
                            <div class="label full section">
                                <div class="tabla-wrapper">
                                    <table class="tabla tabla-verde">
                                        <thead>
                                            <tr>
                                                <th class="sticky" style="width: 70%; text-align: center;">Objetos</th>
                                                <th class="sticky" style="width: 30%; text-align: center;">&nbsp;</th>
                                            </tr>
                                        </thead>  

                                        <tbody>                              
                                            <%
                                                sqlString = "SELECT Secuencia, Descripcion, Nombre, Extension " & _
                                                              "FROM con_Contactos_Adjuntos " & _
                                                             "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                                               "AND (Codigo = '" & Codigo & "');"

                                                set cbox = con.execute(sqlString)
                                                    if not (cbox.bof or cbox.eof) then
                                                        Cuantos = 0

                                                        Do
                                                            Cuantos = Cuantos + 1
                                                            NomObjeto = cbox("Nombre") & "." & cbox("Extension")
                                                            oVinculo = lcase(request.Cookies("usuPath") & "/adjuntos/" & NomObjeto)                                                                   

                                                            %>
                                                                <tr>
                                                                    <td style="width: 20%;"> <%= cbox("Descripcion") %></td>
                                                                    <td style="width: 10%; text-align: center;"> 
                                                                        <button class="form-btn verde" 
                                                                                type="button"
                                                                                onclick="irA('<%= oVinculo %>')">
                                                                            <i class="fa fa-file fa-xl"></i>
                                                                        </button> 

                                                                        <a href="<%= oVinculo %>" download>
                                                                            <button type="button" class="form-btn azul">
                                                                            <i class="fa fa-download fa-xl"></i>
                                                                            </button>                                 
                                                                        </a> 

                                                                        <button class="form-btn rojo" 
                                                                                type="button"
                                                                                onclick="BorrarObjeto('<%= cbox("Secuencia") %>')">
                                                                            <i class="fa fa-trash fa-xl"></i>
                                                                        </button> 
                                                                    </td>
                                                                </tr>
                                                            <%

                                                            cbox.MoveNext
                                                        Loop Until cbox.eof
                                                    end if
                                                cbox.close: set cbox = nothing
                                            %>
                                        </tbody>

                                        <tfoot>
                                            <tr>
                                                <td colspan="4" class="sticky" style="text-align: center;">
                                                    <%
                                                        select case Cuantos
                                                            case 0: response.write "El Contacto No Tiene Adjuntos"
                                                            case 1: response.write "El Contacto Tiene Un Adjunto"
                                                            case else
                                                                response.write "El Contacto Tiene " & Cuantos & " Adjuntos"
                                                        end select
                                                    %>
                                                </td>
                                            </tr>
                                        </tfoot>
                                    </table>
                                </div>

                                <br /><br />

                                <form id="frm_adjuntos" name="frm_adjuntos" action="cont_upload.asp" method="post" enctype="multipart/form-data"> 
                                    <div class="no-ver">
                                        <input id="NuevoObjetoNombre"     name="NuevoObjetoNombre"      type="text"  value="" >
                                        <input id="NuevoObjetoCodigoCont" name="NuevoObjetoCodigoCont"  type="text"  value="<%= Codigo %>" >
                                        <input id="NuevoObjetoCSecuencia" name="NuevoObjetoCSecuencia"  type="text"  value="<%= SecObjeto(Request.Cookies("Usuario"), Codigo) %>" >
                                    </div>

                                    <div class="label full">Descripción del Adjunto</div>

                                    <input class="field" style="width: 100%;" 
                                           id="NuevoObjeto" name="NuevoObjeto" 
                                           type="text" 
                                           placeholder="Descripcion de Nuevo Objeto" >
                                           
                                    <br /><br />

                                    <input type="file" id="File1" name="FILE1" style="width: 100%;" /> 
                                    
                                    <br /><br />

                                    <button class="form-btn verde normal"
                                            id="btn_Adjuntos_Submit" name="btn_Adjuntos_Submit"
                                            onclick="EnviarArchivo()"  type="button"  style="width: 275px; font-size: 16px; color: white;">
                                        <i class="fa fa-folder-open fa-xl"></i>
                                        &nbsp;&nbsp;Subir Archivo&nbsp;&nbsp;
                                    </button>
                                </form>                                
                            </div>
                        </div>
                    <%     
                end sub     

                sub tab_Calendario()
                    dim total                

                    %>
                        <div class="line label-top">
                            <label class="label tiny2" style="vertical-align: top;">Eventos</label>
                            <div class="label full section">
                                <div class="tabla-wrapper">
                                    <table class="tabla tabla-verde">
                                        <thead>
                                            <tr>
                                                <th class="sticky" style="width: 10%; text-align: center;">Fecha</th>
                                                <th class="sticky" style="width: 10%; text-align: center;">Hora</th>
                                                <th class="sticky" style="width: 65%; text-align: center;">Evento</th>
                                                <th class="sticky" style="width: 15%; text-align: center;">Monto</th>
                                            </tr>
                                        </thead>  

                                        <tbody>                              
                                            <%
                                                sqlString = "exec dbo.cont_EventosDbCr '" & Request.Cookies("Usuario") & "', '" & Codigo & "'"

                                                set cbox = con.execute(sqlString)
                                                    if not (cbox.bof or cbox.eof) then
                                                        Cuantos = 0
                                                        total = 0

                                                        Do
                                                            Cuantos = Cuantos + 1
                                                            total = total + cbox("Monto")

                                                            %>
                                                                <tr>
                                                                    <td style="text-align: center;"><%= FechaFormulario(cbox("Fecha")) %></td>
                                                                    <td style="text-align: center;"><%= HoraFormulario(cbox("Hora")) %></td>
                                                                    <td><%= cbox("Descripcion") %></td>
                                                                    <td style="text-align: right;"><%= FormatNumber(cbox("Monto")) %></td>
                                                                </tr>
                                                            <%

                                                            cbox.MoveNext
                                                        Loop Until cbox.eof
                                                    end if
                                                cbox.close: set cbox = nothing
                                            %>
                                        </tbody>

                                        <tfoot>
                                            <tr>
                                                <td colspan="4" class="sticky" style="text-align: right;">
                                                    <%= "Total:&nbsp;&nbsp;" & FormatNumber(total) & "&nbsp;" %>
                                                </td>
                                            </tr>
                                        </tfoot>
                                    </table> 
                                </div>
                            </div>
                        </div>
                    <%     
                end sub  

                sub tab_Notas()
                    %>
                        <div class="line label-top">
                            <label class="label tiny2" style="vertical-align: top;">Notas</label>
                            <div class="label full section">
                                <script src="/core/lib/tinymce/tinymce.min.js"></script>
                                <textarea class="field large editorNotas" name="txtAreaNotas" id="txtAreaNotas"><%= notas %></textarea>

                                <script>
                                    tinymce.init({
                                    entity_encoding : "raw",
                                    selector: '.editorNotas',
                                    license_key: 'gpl',
                                    height: (window.innerHeight - 280),
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
                            </div>
                        </div>
                    <%     
                end sub  

                sub tab_Zodiaco()
                    %>
                        <div class="line label-top">
                            <label class="label tiny2" style="vertical-align: top;">Zoodiaco</label>
                            <div class="label full section">
                                <div class="line-group">
                                    <%
                                        sqlString = "SELECT Signo, Nombre, Desde, Hasta, Naturaleza, Regencia, Fisico, Preferencias, Imagen " & _
                                                    "FROM con_Signos " & _
                                                    "WHERE Signo = " & signo & ";"

                                        set cbox = con.execute(sqlString)
                                            if not (cbox.bof or cbox.eof) then
                                                %>
                                                    <div class="line">
                                                        <label class="label small">Signo</label>
                                                        <input class="field normal" type="text" value="<%= cbox("Nombre") %>" readonly>
                                                    </div>

                                                    <div class="line">
                                                        <label class="label small">Fecha</label>
                                                        <input class="field large" type="text" value="<%= cbox("Desde") & " al " & cbox("Hasta") %>" readonly>
                                                    </div>

                                                    <div class="line">
                                                        <label class="label small">Naturaleza</label>
                                                        <input class="field large" type="text" value="<%= cbox("Naturaleza") %>" readonly>
                                                    </div>

                                                    <div class="line">
                                                        <label class="label small">Regencia</label>
                                                        <input class="field normal" type="text" value="<%= cbox("Regencia") %>" readonly>
                                                    </div>

                                                    <div class="line">
                                                        <label class="label small">Físico</label>
                                                        <input class="field xl" type="text" value="<%= cbox("Fisico") %>" readonly>
                                                    </div>

                                                    <div class="line">
                                                        <label class="label small">Preferencias</label>
                                                        <input class="field xl" type="text" value="<%= cbox("Preferencias") %>" readonly>
                                                    </div>
                                                <%
                                            end if
                                        cbox.close: set cbox = nothing
                                    %>
                                </div>
                            </div>
                        </div>
                    <%     
                end sub   

                sub tab_Arbol()
                    %>
                        <div class="line label-top">
                            <label class="label tiny2">Arbol</label>
                            <div class="label full section">
                                <div class="line-group">
                                    <%
                                        sqlString = "SELECT Arbol, Nombre, Explicacion, Desde, Hasta, Definicion " & _
                                                    "FROM con_Arboles " & _
                                                    "WHERE Arbol = " & arbol & ";"

                                        set cbox = con.execute(sqlString)
                                            if not (cbox.bof or cbox.eof) then
                                                %>
                                                    <div class="line">
                                                        <label class="label small">Nombre</label>
                                                        <input class="field normal" type="text" value="<%= cbox("Nombre") %>" readonly>
                                                    </div>

                                                    <div class="line">
                                                        <label class="label small">Fecha</label>
                                                        <input class="field large" type="text" value="<%= cbox("Desde") & " al " & cbox("Hasta") %>" readonly>
                                                    </div>

                                                    <div class="line">
                                                        <label class="label small">Explicación</label>
                                                        <input class="field large" type="text" value="<%= cbox("Explicacion") %>" readonly>
                                                    </div>

                                                    <div class="line label-top">
                                                        <label class="label small">Definición</label>
                                                        <textarea class="field" style="width: 90%;"  rows="10" cols = "80" type="text" readonly><%= cbox("Definicion") %></textarea>
                                                    </div>
                                                <%
                                            end if
                                        cbox.close: set cbox = nothing
                                    %>
                                </div>
                            </div>
                        </div>
                    <%     
                end sub

                sub tab_Personalidad()
                    %>
                        <div class="line label-top">
                            <label class="label tiny2">Carácter</label>
                            <div class="label full section">
                                <div class="line-group">
                                    <%
                                        sqlString = "SELECT Virtudes, Defectos " & _
                                                    "FROM con_Contactos_Signos " & _
                                                    "WHERE Signo = " & signo & ";"

                                        set cbox = con.execute(sqlString)
                                            if not (cbox.bof or cbox.eof) then
                                                %>
                                                    <div class="line label-top">
                                                        <label class="label small">Virtudes</label>
                                                        <textarea class="field" style="width: 90%;"  rows="10" cols = "80" type="text" readonly><%= cbox("Virtudes") %></textarea>
                                                    </div>

                                                    <div class="line label-top">
                                                        <label class="label small">Defectos</label>
                                                        <textarea class="field" style="width: 90%;"  rows="10" cols = "80" type="text" readonly><%= cbox("Defectos") %></textarea>
                                                    </div>
                                                <%
                                            end if
                                        cbox.close: set cbox = nothing
                                    %>
                                </div>
                            </div>
                        </div>
                    <%     
                end sub  

                sub tab_Relacion()
                    %>
                        <div class="line label-top">
                            <label class="label tiny2">Relación</label>
                            <div class="label full section">
                                <div class="line-group">
                                    <%
                                        sqlString = "SELECT ParaEnamorarle, ParaDejarle " & _
                                                    "FROM con_Contactos_Signos " & _
                                                    "WHERE Signo = " & signo & ";"

                                        set cbox = con.execute(sqlString)
                                            if not (cbox.bof or cbox.eof) then
                                                %>
                                                    <div class="line label-top">
                                                        <label class="label small">Enamorarle</label>
                                                        <textarea class="field" style="width: 90%;"  rows="10" cols = "80" type="text" readonly><%= cbox("ParaEnamorarle") %></textarea>
                                                    </div>

                                                    <div class="line label-top">
                                                        <label class="label small">Separarse</label>
                                                        <textarea class="field" style="width: 90%;"  rows="10" cols = "80" type="text" readonly><%= cbox("ParaDejarle") %></textarea>
                                                    </div>
                                                <%
                                            end if
                                        cbox.close: set cbox = nothing
                                    %>
                                </div>
                            </div>
                        </div>
                    <%     
                end sub

                sub tab_Compatibilidad()
                    %>
                        <div class="line label-top">
                            <label class="label tiny2">Afinidad</label>
                            <div class="label full section">
                                <div class="line-group">
                                    <%
                                        sqlString = "SELECT AfinidadAmorosa, AfinidadLaboral " & _
                                                    "FROM con_Contactos_Afinidades " & _
                                                    "WHERE (signo1 = " & SignoUsuario() & ") " & _
                                                    "AND (signo2 = " & signo & ");"

                                        set cbox = con.execute(sqlString)
                                            if not (cbox.bof or cbox.eof) then
                                                %>
                                                    <div class="line label-top">
                                                        <label class="label small">Laboral</label>
                                                        <textarea class="field" style="width: 90%;"  rows="10" cols = "80" type="text" readonly><%= cbox("AfinidadLaboral") %></textarea>
                                                    </div>

                                                    <div class="line label-top">
                                                        <label class="label small">Sentimental</label>
                                                        <textarea class="field" style="width: 90%;" rows="10" cols = "80" type="text" readonly><%= cbox("AfinidadAmorosa") %></textarea>
                                                    </div>
                                                <%
                                            end if
                                        cbox.close: set cbox = nothing
                                    %>
                                </div>
                            </div>
                        </div>
                    <%     
                end sub     
            '-- Fin: Secciones --
        %>  

        <style>
            .img-limitada {
                display: block;
                margin: 20px auto;
                width: 80%;
                height: auto;
                max-height: 400px;
                min-height: 50px;
                object-fit: contain;
            }

            /* Extendemos y Adaptamos las clases del Framework */

            .campo {
                padding: 0.3rem 0.4rem;
                border: 1px solid #ccc;
                border-radius: 0.3rem;
                font-family: 'Ruda', sans-serif;
                font-size: 1rem;
                color: rgb(25, 25, 25);
                box-sizing: border-box;
                resize: vertical;
            }            
            
            .label.tiny2        { width: 120px ; }
            .field.año          { width: 75px ; }
            .field.description  { width: 550px; } 
        </style>
    </head>

    <body plantilla="normal" reserva="200" onload="verTab(1)">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <br />
        
        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 70%; text-align: left; font-size: 18px; color: rgb(50, 50, 50);">
                <%
                    if (nuevo = 1) then
                        response.write "Contacto Nuevo"
                    else
                        '
                        ' Si (Estatus = 0) apaece el icono de "desborrar"
                        ' Esto cambia el Estatus = 1 y el Visible = 1
                        '

                        if (estatus = 0) then
                            response.write "<img src='/core/imagenes/no_trash.png' height='20' onclick='undelete()'>&nbsp;"
                        else
                            '
                            ' Si (Estatus = 1) entonces 
                            '    Si (Visible = 0) aparce el icono de "Hacer Visible"
                            ' fin si
                            '                            
                            if (visible = 0) then
                                response.write "<img src='/core/imagenes/no_eye.png' height='20' onclick='activate()'>&nbsp;"
                            end if
                        end if

                        response.write nomContacto
                    end if
                %>

                <br />

                <span style="font-size: 16px;">
                    <%
                        if (nuevo = 1) then
                            response.write "&nbsp;"
                        else
                            response.write Categs(Tipo, Codigo)
                        end if
                    %>
                </span>                
            </div>
            
            <div style="flex: 0 0 30%; text-align: right;">
                <button type="button" class="form-btn verde normal" onclick="enviar()" >
                    Grabar
                </button>    

                <button type="button" class="form-btn azul normal" onclick="Volver()" >
                    Cancelar
                </button>                                   
            </div>
        </div> 

        <br />

        <% 
            if FechaCumple = "" then 
                framework_tabs Array("Generales", "Tipo", "Contactos", "Adjuntos", "Eventos", "Notas")
            else
                framework_tabs Array("Generales", "Tipo", "Contactos", "Adjuntos", "Eventos", "Notas", "Zodiaco", "Arbol", "Caracter", "Relacion", "Afinidad")
            end if
        %>

        <br/>

        <div class="main main-scroll">
            <form name="form_transaccion" id="form_transaccion" method="post" action="cont_grabar.asp">
                <div class="no-ver">
                    <input type="text" id="Paquete" name="Paquete" value="<%= Paquete %>">
                </div>

                <div id="tab_Generales" style="display: block;">
                    <table class="tabla-transparente" style="width:100%;">
                        <tr>
                            <td style="width: 60%; vertical-align: top;">
                                <div class="line-group">
                                    <% 
                                        tab_Generales 
                                        response.write "<br />"
                                    %>
                                </div>                          
                            </td>

                            <td style="width:40%; vertical-align: top; text-align: center; border:5px solid red;">
                                <%
                                    if (nuevo = 1) then
                                        fotoObjeto = "/core/imagenes/misc/foto.jpg"
                                    else
                                        fotoObjeto = request.Cookies("usuPath") & "/fotos/" & Codigo & ".jpg"
                                    end if
                                %>

                                <img class="img-limitada" 
                                    src="<%= fotoObjeto %>" 
                                    onerror="this.src='/core/imagenes/misc/foto.jpg'" 
                                    onclick="NuevaFoto()" >
                            </td>
                        </tr>
                    </table><br />

                    <% tab_Telefonos %><br />
                    <% tab_Direcciones %>
                </div>
                
                <div id="tab_Notas" style="display: none;"><% tab_Notas %></div>
            </form>

            <div id="tab_Tipo"      style="display: none;"><% tab_Categorias     %></div>
            <div id="tab_Eventos"   style="display: none;"><% tab_Calendario     %></div>
            <div id="tab_Adjuntos"  style="display: none;"><% tab_Adjuntos       %></div>
            <div id="tab_Zodiaco"   style="display: none;"><% tab_Zodiaco        %></div>
            <div id="tab_Arbol"     style="display: none;"><% tab_Arbol          %></div>
            <div id="tab_Caracter"  style="display: none;"><% tab_Personalidad   %></div>
            <div id="tab_Relacion"  style="display: none;"><% tab_Relacion       %></div>
            <div id="tab_Afinidad"  style="display: none;"><% tab_Compatibilidad %></div>

            <div id="tab_Contactos" style="display: none;">
                <% 
                    tab_ContactosRelacionados
                    response.write "<br />"
                    tab_ContactosNoRelacionados
                %>
            </div>
        </div>    

        <br /><br />

        <script type="text/javascript">
            function irA(vinculo) {
                window.location.href = vinculo;
            }

            function enviar() {
                document.getElementById("form_transaccion").submit();
            }

            function Volver() {
                var vinculo = "lista.asp?v=<%= ver %>&t=<%= tipo %>&c=<%= categ %>&o1=<%= orden1 %>&o2=<%= orden2 %>";
                window.location.href = vinculo;
            }            

            function NuevoTel() {
                var cod = document.getElementById("cod").value;
                var tel = document.getElementById("nuevoTelefono").value;
                var tipo = document.getElementById("nuevoTipo").value;

                tel = tel.replace("+","*");

                if (tel != "") {
                    var vinculo = "cont_nuevo_telefono.asp?c=" + cod + "&t=" + tel + "&l=" + tipo + "&tt=1";

                    window.location.href = vinculo;
                }
            }

            function BorrarTel(Secuencia, Telefono) {
                var confirmacion = confirm("Está seguro de borrar el telefono '" + Telefono + "'?");

                var cod = document.getElementById("cod").value;
                vinculo = "cont_borrar_telefono.asp?c=" + cod + "&s=" + Secuencia + "&tt=1";

                if (confirmacion) {
                    window.location.href = vinculo;
                } else {
                    window.alert("Proceso Cancelado.");        
                } 
            }

            function NuevaCat(tipo) {
                var cod = document.getElementById("cod").value;          
                var categ = document.getElementById("frm_NuevaCateg").value;

                if (categ == "*") {
                    alert("Debe seleccionar un ítem de la lista.");
                } else {
                    var vinculo = "cont_nueva_categ.asp?t=" + tipo + "&c=" + cod + "&k=" + categ + "&tt=3";
                    window.location.href = vinculo;
                }
            }

            function BorrarCat(Contacto, Tipo, Categoria, NombreCategoria) {
                var confirmacion = confirm("Está seguro de borrar la Categoria '" + NombreCategoria + "'?");

                var cod = document.getElementById("cod").value;
                vinculo = "cont_borrar_categ.asp?t=" + Tipo + "&c=" + Contacto + "&k=" + Categoria + "&tt=3";

                if (confirmacion) {
                    window.location.href = vinculo;
                } else {
                    window.alert("Proceso Cancelado.");        
                } 
            }    

            function NuevoContRelacionado() {
                var cod = document.getElementById("cod").value;          
                var rel = document.getElementById("NuevaRelacion").value;
                var cont = document.getElementById("NuevoContactoRelacionado").value;
                var cumple = document.getElementById("NuevoCumpleContactoRelacionado").value;

                var vinculo = "cont_nuevo_cont_rel.asp?c=" + cod + "&r=" + rel + "&k=" + cont + "&q=" + cumple + "&tt=3";
                window.location.href = vinculo;
            }    

            function BorrarContRelacionado(Secuencia, Nombre) {
                var confirmacion = confirm("Está seguro de borrar el Contacto '" + Nombre + "'?");

                var cod = document.getElementById("cod").value;         
                vinculo = "cont_borrar_cont_rel.asp?c=" + cod + "&s=" + Secuencia + "&tt=3";

                if (confirmacion) {
                    window.location.href = vinculo;
                } else {
                    window.alert("Proceso Cancelado.");        
                } 
            }       

            function NuevoContNoRelacionado() {
                var cod = document.getElementById("cod").value;          
                var rel = document.getElementById("NuevaNoRel").value;
                var cont = document.getElementById("NuevoContNoRel").value;
                var cumple = document.getElementById("NuevoCumpleContNoRel").value;

                var vinculo = "cont_nuevo_cont_no_rel.asp?c=" + cod + "&r=" + rel + "&k=" + cont + "&q=" + cumple + "&tt=3";
                window.location.href = vinculo;
            }      

            function BorrarContNoRelacionado(Secuencia, Nombre) {
                var confirmacion = confirm("Está seguro de borrar el Contacto '" + Nombre + "'?");

                var cod = document.getElementById("cod").value;         
                vinculo = "cont_borrar_cont_no_rel.asp?c=" + cod + "&s=" + Secuencia + "&tt=3";

                if (confirmacion) {
                    window.location.href = vinculo;
                } else {
                    window.alert("Proceso Cancelado.");        
                } 
            }  

            function EnviarArchivo(){
                var fileInput = document.getElementById("File1");
                var fileName = fileInput.value.split(/(\\|\/)/g).pop();

                document.getElementById("NuevoObjetoNombre").value = fileName;
                document.getElementById("frm_adjuntos").submit();
            }

            function BorrarObjeto(Secuencia) {
                var confirmacion = confirm("Está seguro de borrar este adjunto?");

                var cod = document.getElementById("cod").value;         
                vinculo = "cont_borrar_adjunto.asp?s=" + Secuencia + "&tt=4";

                if (confirmacion) {
                    window.location.href = vinculo;
                } else {
                    window.alert("Proceso Cancelado.");        
                } 
            }  

            function NuevaFoto() {
                var vinculo = "cont_foto.asp?con=<%= Codigo %>&v=<%= ver %>&t=<%= tipo %>&c=<%= categ %>&o1=<%= orden1 %>&o2=<%= orden2 %>";
                window.location.href = vinculo;          
            } 

            function undelete() {
                var confirmacion = confirm("Quiere recuperar este contacto y hacerlo visible nuevamente?");

                if (confirmacion) {
                    var vinculo = "cont_undelete.asp?con=<%= Codigo %>&v=<%= ver %>&t=<%= tipo %>&c=<%= categ %>&o1=<%= orden1 %>&o2=<%= orden2 %>";
                    window.location.href = vinculo;
                } else {
                    window.alert("Proceso Cancelado.");        
                }
            }

            function activate() {
                var confirmacion = confirm("Volver a hacer visible a este contacto?");

                if (confirmacion) {
                    var vinculo = "cont_activate.asp?con=<%= Codigo %>&v=<%= ver %>&t=<%= tipo %>&c=<%= categ %>&o1=<%= orden1 %>&o2=<%= orden2 %>";
                    window.location.href = vinculo;
                } else {
                    window.alert("Proceso Cancelado.");        
                }
            }        
            
            mask(document.getElementById('fechaCumple'),                      ['99/99']);
            mask(document.getElementById('NuevoCumpleContactoRelacionado'),   ['99/99']);
            mask(document.getElementById('NuevoCumpleContNoRel'),             ['99/99']);
        </script>

        <% con.close: set con = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->

    </body>
</html>