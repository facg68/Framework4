<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->   
        <%
            thisSystem = "discos"
            thisProcess = "discos.0110"
            SysLockOut


            '
            ' Init()
            '

            dim con, t, tt, sqlString
            dim ACompra, AEdicion, Titulo, Precio, Tienda, Casa, Descripcion, VerComo, Carpeta 

            Usuario = Request.Cookies("usuario")
            Paquete = Request.QueryString("m")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            sqlString = "select ACompra, AEdicion, Titulo, Precio, Tienda, Casa, Descripcion, VerComo, Carpeta " & _
                        "from discos_Paquetes " & _
                        "WHERE (Usuario = '" & Usuario & "') " & _
                        "AND (Paquete = '" & Paquete & "');"

            set t = con.execute(sqlString)
                ACompra = t("ACompra")
                AEdicion = t("AEdicion")
                Titulo = t("Titulo")
                Precio = t("Precio")
                Tienda = t("Tienda")
                Casa = t("Casa")
                Descripcion = t("Descripcion")
                VerComo = t("VerComo")
                Carpeta = t("Carpeta")
            t.close: set t = nothing 
        %>  

        <style>
            .line-group .line {
                margin-bottom: 12px;
            }

            .line-group .line:last-child {
                margin-bottom: 0;
            }      

            .img-limitada {
                display: block;
                margin: 0 auto;
                width: 100%;
                height: auto;
                max-height: 500px;
                min-height: 50px;
                object-fit: contain;
            }

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

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <br />

        
        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 70%; text-align: left; font-size: 18px; color: rgb(50, 50, 50);">
                    (<%= Paquete %>) - <%= Titulo %>
            </div>
            
            <div style="flex: 0 0 30%; text-align: right;">
                <button type="button" class="form-btn verde normal" onclick="enviar()" >
                    Grabar
                </button>    

                <button type="button" class="form-btn azul normal" onclick="irA('lista.asp')" >
                    Cancelar
                </button>                        
            </div>
        </div>        

        <div class="main main-scroll">           
            <form name="form_transaccion" id="form_transaccion" method="post" action="grabar_registro.asp">
                <!--
                    Parte 1: Generales del Paquete
                -->

                <div class="no-ver">
                    <input type="text" id="Paquete" name="Paquete" value="<%= Paquete %>">
                </div>

                <table class="tabla-transparente" style="width:100%; table-layout: fixed;">
                    <tr>
                        <td style="width:65%;">
                            <div class="line-group">
                                <div class="line">
                                    <label class="label tiny2" onclick="Filtro('Amo = <%= AEdicion %>', 'Año Editado = <%= AEdicion %>')" >Año Edición</label>
                                    <input class="field año" type="text" id="AEdicion" name="AEdicion" value="<%= AEdicion %>" placeholder="9999" required>
                                </div>

                                <div class="line">
                                    <label class="label tiny2" onclick="Filtro('AmoCompra = <%= ACompra %>', 'Año Compra = <%= ACompra %>')" >Año Compra</label>
                                    <input class="field año" type="text" id="ACompra" name="ACompra" value="<%= ACompra %>" placeholder="9999" required>
                                </div>

                                <div class="line">
                                    <label class="label tiny2">Título</label>
                                    <input class="field description" type="text" id="Titulo" name="Titulo" value="<%= Titulo %>" placeholder="Título del Paquete" required>
                                </div>

                                <div class="line">
                                    <label class="label tiny2" onclick="Filtro('CasaDisquera = <%= Casa %>', 'Casa Disquera = <%= Casa %>')" >Casa Editora</label>
                                    <%
                                        sqlString = "SELECT Codigo, Nombre " & _
                                                    "FROM discos_Casas " & _ 
                                                    "WHERE Usuario = '" & Usuario & "' " & _
                                                    "ORDER BY Nombre ASC;"

                                        set tt = con.execute(sqlString)
                                            if not (tt.bof or tt.eof) then
                                                response.write "<select class='field description' name='Casa' id='Casa' required >" 
                                                    Do
                                                        response.write "<option value='" & tt("Codigo") & "' "
                                                            if Casa = tt("Codigo") then response.write " selected" 
                                                        response.write ">" & tt("Nombre") & "</option>"

                                                        tt.MoveNext
                                                    Loop Until tt.eof
                                                response.write "</select>"
                                            end if
                                        tt.close: set tt = nothing
                                    %>  
                                </div>

                                <div class="line">
                                    <label class="label tiny2" onclick="Filtro('Tienda = <%= Tienda %>', 'Tienda = <%= Tienda %>')" >Tienda</label>
                                    <%
                                        sqlString = "SELECT Codigo, Nombre " & _
                                                    "FROM discos_Tiendas " & _ 
                                                    "WHERE Usuario = '" & Usuario & "' " & _
                                                    "ORDER BY Nombre ASC;"

                                        set tt = con.execute(sqlString)
                                            if not (tt.bof or tt.eof) then
                                                response.write "<select class='field description'name='Tienda' id='Tienda' required >" 
                                                    Do
                                                        response.write "<option value='" & tt("Codigo") & "' "
                                                            if Tienda = tt("Codigo") then response.write " selected" 
                                                        response.write ">" & tt("Nombre") & "</option>"

                                                        tt.MoveNext
                                                    Loop Until tt.eof
                                                response.write "</select>"
                                            end if
                                        tt.close: set tt = nothing
                                    %>   
                                </div>    

                                <div class="line">
                                    <label class="label tiny2">Precio</label>
                                    <input class="field tiny" type="text"  id="Precio" name="Precio" value="<%= Precio %>" placeholder="00.00" required>
                                </div>  

                                <div class="line label-top">
                                    <label class="label tiny2">Descripción</label>
                                    <textarea class="field description" 
                                            name="Descripcion" id="Descripcion" 
                                            rows=5 cols=80
                                            style="font-family: courier; font-size: 16px;"><%= Descripcion %></textarea>
                                </div> 

                                <div class="line">
                                    <%
                                        Select Case VerComo
                                            Case 1: vc_Titulo = "Tipo de Vista = Ver Como Paquete"
                                            Case 2: vc_Titulo = "Tipo de Vista = Ver el Contenido (Ver Objetos)"
                                            Case 0: vc_Titulo = "Tipo de Vista = Ocultar Paquete (No Ver)"
                                        End Select

                                        vc_Filtro = "VerComo = " & VerComo
                                    %>

                                    <label class="label tiny2" onclick="Filtro('<%= vc_Filtro %>', '<%= vc_Titulo %>')">Ver Como</label>
                                    <select class="field small" name="VerComo" id="VerComo" required >
                                        <option value="1" <% if VerComo = "1" then response.write " selected" %>>Paquete</option>
                                        <option value="2" <% if VerComo = "2" then response.write " selected" %>>Objetos</option>
                                        <option value="0" <% if VerComo = "0" then response.write " selected" %>>No Ver</option>             
                                    </select>    
                                </div>  

                                <div class="line">
                                    <label class="label tiny2" onclick="Filtro('Coleccion = <%= Carpeta %>', 'Colecci&oacute;n = <%= Carpeta %>')" >Colección</label>
                                    <%
                                        sqlString = "SELECT Codigo, Nombre " & _
                                                    "FROM discos_Carpetas " & _ 
                                                    "WHERE Usuario = '" & Usuario & "' " & _
                                                    "ORDER BY Nombre ASC;"

                                        set tt = con.execute(sqlString)
                                            if not (tt.bof or tt.eof) then
                                                response.write "<select class='field description' name='Carpeta' id='Carpeta' required >" 
                                                    Do
                                                        response.write "<option value='" & tt("Codigo") & "' "
                                                            if Carpeta = tt("Codigo") then response.write " selected" 
                                                        response.write ">" & tt("Nombre") & "</option>"

                                                        tt.MoveNext
                                                    Loop Until tt.eof
                                                response.write "</select>"
                                            end if
                                        tt.close: set tt = nothing
                                    %>     
                                </div>  
                            </div>                          
                        </td>

                        <td style="width:35%; vertical-align: top; text-align: center;">
                            <%
                                if (nuevo = 1) then
                                    response.write "<img src='/core/imagenes/misc/foto.jpg' alt='Portada del Medio'>"
                                else
                                    fotoPath = request.Cookies("usuPath") & "/medios/" & Paquete & ".jpg"
                                    %>
                                    <img class="img-limitada" 
                                            name="portada" id="portada" src="<%= fotoPath %>" 
                                            onerror="this.src='/core/imagenes/misc/foto.jpg'" 
                                            onclick="NuevaFoto('<%= Paquete %>')">
                                    <%
                                end if
                            %> 
                        </td>
                    </tr>
                </table>

                <!--
                    Parte 2: Sección - Objetos del Paquete
                -->                    

                <div class="line label-top">
                    <label class="label tiny2" style="vertical-align: top;">Objetos:</label>
                    <div class="label full section">
                        <div class="tabla-wrapper">
                            <table class="tabla tabla-blue">
                                <thead>
                                    <tr>
                                        <th colspan="4" class="sticky">Lista de Objetos</th>
                                        <th colspan="1" class="sticky" style="text-align: right;">
                                            <select class="campo" style="background-color: #f0faff; width: 100%;"
                                                    name="cboNuevoTipoObjeto" id="cboNuevoTipoObjeto"
                                                    onchange="NuevoObjeto()" >
                                                <option value="*" selected>- - Nuevo - -</option>
                                                <%
                                                    set tt = con.execute("select Codigo, Nombre from discos_Objetos_Clases ORDER BY Nombre ")

                                                    if not (tt.bof or tt.eof) then
                                                        Do
                                                            response.write "<option value='" & tt("Codigo") & "'"
                                                            if Tipo = tt("Codigo") then response.write " selected" 
                                                            response.write ">" & tt("Nombre") & "</option>"

                                                            tt.MoveNext
                                                        Loop Until tt.eof
                                                    end if

                                                    tt.close: set tt = nothing
                                                %>
                                            </select>
                                        </th>
                                    </tr>
                                </thead>

                                <tbody>
                                    <%
                                        sqlString = "select Paquete, Objeto, AEdicion, Titulo, InDirAu, Forma, Es3D, Genero, PlatOS, Plataforma, Editor, Visible, Icono_Forma " & _
                                                    "from discos_ObjetosPaquete " & _
                                                    "where (Usuario = '" & Usuario & "') " & _
                                                    "AND (Paquete = '" & Paquete & "') "   
                                        cuantos = 0

                                        set tt = con.execute(sqlString)

                                        if not (tt.bof or tt.eof) then
                                            Do
                                                vinculo = "editar_objeto.asp?p=" & Paquete & "&o=" & tt("Objeto") & "&e=" & tt("Editor")
                                                fotoPaquete = "/perfiles/" & Usuario & "/medios/" & tt("Paquete") & "_s.jpg"
                                                fotoObjeto = "/perfiles/" & Usuario & "/medios/" & tt("Objeto") & "_s.jpg"
                                                icn_Forma = "/perfiles/" & Usuario & "/discos/" & tt("Icono_Forma")
                                                icn_3D= "/perfiles/" & Usuario & "/discos/icono_3D.gif"

                                                cuantos = cuantos + 1

                                                %>
                                                    <tr>
                                                        <td style="width: 10%; text-align: center;" onclick="irA('<%= vinculo %>')">
                                                            <img src="<%= fotoObjeto %>" onerror="this.src='<%= fotoPaquete %>'" width="50">
                                                        </td>

                                                        <td style="width: 10%; text-align: center;" onclick="irA('<%= vinculo %>')">
                                                            <img src="<%= icn_Forma %>" width="50">
                                                        </td>    

                                                        <td style="width: 46%;" onclick="irA('<%= vinculo %>')">
                                                            <span style="font-family: 'Ruda Bold';">
                                                                <%= tt("Titulo") %>
                                                            </span>
                                                            
                                                            <br/>
                                                            
                                                            <%
                                                                if len(trim(tt("InDirAu"))) > 0 then
                                                                    response.write tt("InDirAu") & "<br/>"
                                                                end if

                                                                if tt("PlatOS") <> "00000000" then
                                                                    response.write tt("Plataforma") & ", " 
                                                                else
                                                                    response.write tt("Genero") & ", " 
                                                                end if

                                                                response.write tt("AEdicion")
                                                            %>
                                                        </td>

                                                        <td style="width: 10%; text-align: center;" onclick="irA('<%= vinculo %>')">
                                                            <%
                                                                if tt("Es3D") = 1 then
                                                                    response.write "<img src='" & icn_3D & "' width='50'>"
                                                                else
                                                                    response.write "&nbsp;"
                                                                end if                                                
                                                            %>                                                
                                                        </td>

                                                        <td style="width: 24%; text-align: center;" onclick="irA('<%= vinculo %>')">
                                                            <%
                                                                response.write "<button class='form-btn "
                                                                    if tt("Visible") = 1 then
                                                                        response.write "verde"
                                                                    else
                                                                        response.write "rojo"
                                                                    end if
                                                                response.write "' "%> onClick="visibilidad('<%= Paquete %>', '<%= tt("Objeto") %>')"><%

                                                                response.write "<i class=' fa fa-eye fa-xl' title='"
                                                                    if tt("Visible") = 1 then
                                                                        response.write "Desactivar Visibilidad"
                                                                    else
                                                                        response.write "Activar Visibilidad"
                                                                    end if                    
                                                                response.write "'></i></button>&nbsp;"

                                                                %><button class="form-btn azul" onClick="copiarObjeto('<%= Paquete %>', '<%= tt("Objeto") %>', '<%= tt("Editor") %>')"><%
                                                                response.write "<i class=' fa fa-copy fa-xl' title='Duplicar Objeto'></i>"
                                                                response.write "</button>&nbsp;"

                                                                %><button class="form-btn rojo" onClick="borrarObjeto('<%= Paquete %>', '<%= tt("Objeto") %>', '<%= tt("Editor") %>')"><%
                                                                response.write "<i class=' fa fa-trash fa-xl' title='Borrar Objeto'></i>"
                                                                response.write "</button>"                                             
                                                            %>                                                
                                                        </td>  
                                                    </tr> 
                                                <%                                         
            
                                                tt.MoveNext
                                            Loop Until (tt.eof)
                                        end if

                                        tt.close: set tt = nothing
                                    %> 
                                </tbody>

                                <tfoot>
                                    <tr>
                                        <td colspan="5" class="sticky" style="text-align: center;">
                                            <%
                                                select case cuantos
                                                    case 0: response.write "El Paquete No Tiene Objetos"
                                                    case 1: response.write "El Paquete Tiene Un Objeto"
                                                    case else
                                                        response.write "El Paquete Tiene " & Cuantos & " Objetos"
                                                end select
                                            %>
                                        </td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div> 
                    </div>
                </div>  

                <!--
                    Parte 3: Sección - Metadatos
                -->                    

                <div class="line label-top">
                    <label class="label tiny2" style="vertical-align: top;">Metadata:</label>
                    <div class="label full section">
                        <div class="tabla-wrapper">
                            <table class="tabla tabla-violet">
                                <thead>
                                    <tr>
                                        <th colspan="2" class="sticky">Llaves</th>
                                    </tr>
                                </thead>  

                                <tbody>                              
                                    <%
                                        sqlString = "SELECT MetaData " & _
                                                    "FROM discos_Paquetes_Metadata " & _
                                                    "WHERE (Usuario = '" & Usuario & "') " & _
                                                    "AND (Paquete = '" & Paquete & "') " & _
                                                    "ORDER BY MetaData;"

                                        set cbox = con.execute(sqlString)
                                            if not (cbox.bof or cbox.eof) then
                                                Cuantos = 0

                                                Do
                                                    Cuantos = Cuantos + 1

                                                    %>
                                                        <tr>
                                                            <td style="width: 90%;"><%= cbox("Metadata") %></td>

                                                            <td style="width: 10%; text-align: center;">
                                                                <button class = "form-btn rojo" 
                                                                        type = "button" 
                                                                        onClick = "BorrarMetaData('<%= Paquete %>', '<%= cbox("Metadata") %>')"
                                                                >
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
                                        más metadata...
                                    -->                                    

                                    <tr>
                                        <td style="width: 90%;">
                                            <input class="field" 
                                                    style="width: 100%; background-color: transparent; border: transparent;" 
                                                    type="text" 
                                                    id="NuevaMetaData" name="NuevaMetaData" placeholder="Nueva Llave de Metadata">
                                        </td>

                                        <td style="width: 10%; text-align: center;">
                                            <button class="form-btn verde" type="button" onclick="Nueva_MetaData('<%= Paquete %>')"><i class="fa fa-save fa-xl"></i></button>
                                        </td>
                                    </tr>
                                </tbody>

                                <tfoot>
                                    <tr>
                                        <td colspan="2" class="sticky" style="text-align: center;">
                                            <%
                                                select case Cuantos
                                                    case 0: response.write "El Paquete No Tiene Metadata"
                                                    case 1: response.write "El Paquete Tiene Una Llave"
                                                    case else
                                                        response.write "El Paquete Tiene " & Cuantos & " Llaves"
                                                end select
                                            %>
                                        </td>
                                    </tr>
                                </tfoot>
                            </table> 
                        </div>
                    </div>
                </div>
            </form>
        </div>    

        <br /><br />

        <script type="text/javascript">
            function irA(vinculo) {
                window.location.href = vinculo;
            }

            function enviar() {
                document.getElementById("form_transaccion").submit();
            }

            function NuevoObjeto() {
                var confirmacion = confirm("Desea Añadir Un Objeto?");

                if (confirmacion) {
                    var tipo = document.getElementById("cboNuevoTipoObjeto").value;

                    if (tipo != "*") {
                        document.getElementById("form_transaccion").submit(); 
                    } else {
                        alert("Debe Seleccionar un tipo de Objeto!");
                    }
                } else {
                    document.getElementById("cboNuevoTipoObjeto").value = "*";
                }
            }

            function Nueva_MetaData(paquete) {
                var meta = document.getElementById("NuevaMetaData").value;
                var vinculo = "metadata_grabar.asp?p=" + paquete + "&m=" + meta;
             
                if (meta != "") {
                    window.location.href = vinculo;
                } else {
                    window.alert("El Valor de la MetaData No se ha especificado!");
                };
            }

            function BorrarMetaData(paquete, metadata) {
                var confirmacion = confirm("Desea borrar la MetaData " + metadata + "?");
                var vinculo = "metadata_borrar.asp?p=" + paquete + "&m=" + metadata;

                if (confirmacion) {     
                    window.location.href = vinculo;
                };
            } 

            function NuevaFoto(paquete) {
                var vinculo = "editar_paquete_foto.asp?p=" + paquete;
                window.location.href = vinculo;          
            }  

            function visibilidad(paquete, objeto) {
                var vinculo = "objeto_visibilidad.asp?p=" + paquete + "&o=" + objeto;
                window.location.href = vinculo;
            }

            function copiarObjeto(paquete, objeto, editor) {
                var vinculo = "objeto_duplicar.asp?p=" + paquete + "&o=" + objeto + "&e=" + editor;
                window.location.href = vinculo;        
            }

            function borrarObjeto(paquete, objeto, editor) {
                var confirmacion = confirm("Desea borrar el Objeto seleccionado?");
                var vinculo = "objeto_borrar.asp?p=" + paquete + "&o=" + objeto + "&e=" + editor;

                if (confirmacion) {  
                window.location.href = vinculo;
                };        
            }

            function Filtro(cadena, titulo) {
                vinculo = "/forma/desktop/apps/discos/lista_paquetes.asp?f=" + cadena + "&t=" + titulo;
                window.location.href = vinculo;
            }            

            mask(document.getElementById('AEdicion'), ['9999']);
            mask(document.getElementById('ACompra'),  ['9999']);
            mask(document.getElementById('Precio'),   ['9.99', '99.99', '999.99', '9999.99',]);
        </script>

        <% con.close: set con = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>