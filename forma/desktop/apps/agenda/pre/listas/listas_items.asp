<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <title>Editar Items de las Listas</title>    
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->             
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0110"
            SysLockOut

            '
            ' Declaraciones 
            '

            dim sqlString, tot_Precio, tot_Cambio
            dim Usuario, Codigo, Nombre, Descripcion, Monto, Contacto, PrecioOriginal
            dim PrecioFinal, Grupo, Categoria, VerListaEnInforme

            dim cuantos, nombreLista, MultiPrecio, Cuenta, MonedaOrigen, MonedaDestino, TipoDetalle, numCols
            dim Secuencia, Item, Precio, Fecha, t1, conn, t

            Usuario = Request.Cookies("Usuario")
            Codigo = Request.QueryString("l")
            CargarParametros Codigo

            '
            ' Funciones y Procedimientos
            '            

            function LocalMonetarioUsuario()
                dim cc, f

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")               
                    set f = con.execute("SELECT isnull(usuLocal, 'US') AS usuLocal FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';")
                        LocalMonetarioUsuario = f("usuLocal")               
                    f.close: set f = nothing
                cc.close: set cc = nothing
            end function

            Sub CargarParametros(CodigoLista)
                dim cc, f, ssql

                ssql = "SELECT Nombre, MultiPrecio, Cuenta, PrecioOriginal, PrecioFinal " & _
                       "FROM pre_Listas_Encabezado " & _
                       "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                       "AND (Codigo = '" & CodigoLista & "');"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")                       
                    set f = cc.execute(ssql)
                        NombreLista = f("Nombre")
                        Multiprecio = f("MultiPrecio")
                        Cuenta = f("Cuenta")
                        MonedaOrigen = f("PrecioOriginal")
                        MonedaDestino = f("PrecioFinal")  

                        if Cuenta = 1 then 
                            If Multiprecio = 0 then
                                TipoDetalle = 1
                            else
                                TipoDetalle = 2
                            end if                
                        else
                            If Multiprecio = 0 then
                                TipoDetalle = 3
                            else
                                TipoDetalle = 4
                            end if
                        end if
                    f.close: set f = nothing
                cc.close: set cc = nothing
            End Sub  

            function FechaServer(FechaForm)
                dim d, m, a

                if FechaForm = "" then 
                    FechaServer = NULL
                else
                    d = left(FechaForm, 2)
                    m = mid(FechaForm, 4, 2)
                    a = right(FechaForm, 4)

                    FechaServer = a & "-" & right("00" & m, 2) & "-" & right("00" & d, 2)
                end if
            end function   

            function FechaFormulario(FechaServer)
                dim d, m, a

                if FechaServer <> "" then      
                    d = right("00" & right(FechaServer, 2), 2)
                    m = right("00" & mid(FechaServer, 6, 2), 2)
                    a = left(FechaServer, 4)

                    FechaFormulario = d & "/" & m & "/" & a
                end if
            end function    
        %>    

        <style>
            .encabezado {
                font-size: 0.85rem;
                font-weight: bold;
                color: #444;
                white-space: nowrap;
                padding: 0px;
                padding-left: 10px;                
            }            

            .subformulario tr {
                border-bottom: none !important;
            }
        </style>
    </head>

    <body plantilla="tabla" reserva="245">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            cuantos = 0 
            tot_Precio = 0.00
            tot_Cambio = 0.00   

            set conn = Server.CreateObject("ADODB.Connection")
            conn.open Application("Conn")                                  

            sqlString = "exec pre_Listas_Items '" & Usuario & "', '" & Codigo & "'"       
            set t = conn.execute(sqlString)        
        %>

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 65%; text-align: left; font-size: 20px; color: rgb(50, 50, 50);">
                <input class="field frame" 
                    style="width: 100%; text-align: left;" 
                    type="text"
                    value="Items de <%= NombreLista %>" 
                    readonly
                >                
            </div>
            
            <div style="flex: 0 0 35%; text-align: right;">
                <button type='button' class='form-btn azul' style='width: 100px; font-size: 16px; color: white;' onclick="enviarFormulario()">Actualizar</button>
               
                &nbsp;&nbsp;

                <button type='button' class='form-btn rojo' style='width: 100px; font-size: 16px; color: white;' onclick="abrir('lista.asp')">Cancelar</button>
            </div>
        </div>      

        <div class="main" style="width: 95%;">
            <!-- Lista de Items -->
                <form name="form_lista" id="form_lista" method="post" action="listas_items_actualizar.asp">
                    <div class="no-ver">
                        <input id="Codigo"      name="Codigo"       value="<%= Codigo %>">
                        <input id="Usuario"     name="Usuario"      value="<%= Usuario %>">
                        <input id="MultiPrecio" name="MultiPrecio"  value="<%= MultiPrecio %>">
                    </div>  

                    <div class="line">
                        <div class="tabla-wrapper">
                            <table class="tabla tabla-blue">
                                <thead>
                                    <tr>
                                        <%
                                            if Cuenta = 1 then
                                                if MultiPrecio = 1 then
                                                    filaTipo = 1
                                                    %>
                                                        <th class="sticky" style="width: 60%; text-align: center;">Item</td>
                                                        <th class="sticky" style="width: 15%; text-align: center;">Precio</td>
                                                        <th class="sticky" style="width: 15%; text-align: center;">Cambio</td>
                                                        <th class="sticky" style="width: 10%; text-align: center;">&nbsp;</td>                                                
                                                    <%                
                                                else
                                                    filaTipo = 2
                                                    %>
                                                        <th class="sticky" style="width: 75%; text-align: center;">Item</td>
                                                        <th class="sticky" style="width: 15%; text-align: center;">Precio</td>
                                                        <th class="sticky" style="width: 10%; text-align: center;">&nbsp;</td>                                                
                                                    <%                                              
                                                end if
                                            else
                                                if MultiPrecio = 1 then
                                                    filaTipo = 3
                                                    %>
                                                        <th class="sticky" style="width: 50%; text-align: center;">Item</td>
                                                        <th class="sticky" style="width: 15%; text-align: center;">Precio</td>
                                                        <th class="sticky" style="width: 15%; text-align: center;">Cambio</td>
                                                        <th class="sticky" style="width: 10%; text-align: center;">Fecha</td>
                                                        <th class="sticky" style="width: 10%; text-align: center;">&nbsp;</td>                                                
                                                    <%                                                  
                                                else
                                                    filaTipo = 4
                                                    %>
                                                        <th class="sticky" style="width: 60%; text-align: center;">Item</td>
                                                        <th class="sticky" style="width: 15%; text-align: center;">Precio</td>
                                                        <th class="sticky" style="width: 15%; text-align: center;">Fecha</td>
                                                        <th class="sticky" style="width: 10%; text-align: center;">&nbsp;</td>                                                
                                                    <%                                             
                                                end if
                                            end if
                                        %>
                                    </tr>
                                </thead>

                                <tbody>
                                    <%
                                        if not (t.bof or t.eof) then
                                            do
                                                cuantos = cuantos + 1

                                                nombreItem           = "Litem_" & t("Llave")
                                                nombrePrecioOriginal = "Lpori_" & t("Llave")
                                                nombrePrecio         = "Lprec_" & t("Llave")
                                                nombreFecha          = "Lfech_" & t("Llave")

                                                response.write "<tr>"
                                                    Select Case TipoDetalle
                                                        Case 1
                                                            '
                                                            ' Es una Cuenta. No aparece el campo "Fecha"
                                                            '
                                                            tot_Precio = tot_Precio + cDbl(t("PrecioOriginal"))

                                                            %>
                                                                <td>
                                                                    <input class="field frame" style="width: 100%;" name="<%= nombreItem %>" id="<%= nombreItem %>" type="text" value="<%= t("Item") %>">
                                                                </td>

                                                                <td>
                                                                    <input class="field frame" 
                                                                        style="width: 100%; text-align: right;" 
                                                                        name="<%= nombrePrecioOriginal %>" 
                                                                        id="<%= nombrePrecioOriginal %>" 
                                                                        type="number" step="0.01" placeholder="0.00" 
                                                                        value="<%= t("PrecioOriginal") %>" required
                                                                    >
                                                                </td>
                                                            <%   
                                                        Case 2
                                                            '
                                                            ' Es una Cuenta Multiprecio. No aparece el campo "Fecha" pero aparecen dos precios distintos
                                                            ' 

                                                            tot_Precio = tot_Precio + cDbl(t("PrecioOriginal"))
                                                            tot_Cambio = tot_Cambio + cDbl(t("Precio"))

                                                            %>
                                                                <td>
                                                                    <input class="field frame" style="width: 100%;" name="<%= nombreItem %>" id="<%= nombreItem %>" type="text" value="<%= t("Item") %>">
                                                                </td>

                                                                <td>
                                                                    <input class="field frame" style="width: 100%; text-align: right;" 
                                                                        name="<%= nombrePrecioOriginal %>" 
                                                                        id="<%= nombrePrecioOriginal %>" 
                                                                        type="number" step="0.01" placeholder="0.00" 
                                                                        value="<%= t("PrecioOriginal") %>" 
                                                                        onChange="CambiarMonedaLista('<%= t("LocalOrigen") %>', '<%= t("LocalDestino") %>', 1, '<%= nombrePrecioOriginal %>', '<%= nombrePrecio %>');"
                                                                        required
                                                                    >
                                                                </td>

                                                                <td>
                                                                    <input class="field frame" style="width: 100%; text-align: right;" 
                                                                        name="<%= nombrePrecio %>" 
                                                                        id="<%= nombrePrecio %>" 
                                                                        type="number" step="0.01" placeholder="0.00" 
                                                                        value="<%= t("Precio") %>" 
                                                                        onChange="CambiarMonedaLista('<%= t("LocalDestino") %>', '<%= t("LocalOrigen") %>', 1, '<%= nombrePrecio %>', '<%= nombrePrecioOriginal %>');"
                                                                        required
                                                                    >
                                                                </td>                                         
                                                            <%   
                                                        Case 3
                                                            '
                                                            ' No es una Cuenta. Aparece el campo "Fecha"
                                                            '           

                                                            tot_Precio = tot_Precio + cDbl(t("PrecioOriginal"))

                                                            %>
                                                                <td>
                                                                    <input class="field frame" style="width: 100%;" name="<%= nombreItem %>" id="<%= nombreItem %>" type="text" value="<%= t("Item") %>">
                                                                </td>

                                                                <td>
                                                                    <input class="field frame" style="width: 100%; text-align: right;" 
                                                                        name="<%= nombrePrecioOriginal %>" 
                                                                        id="<%= nombrePrecioOriginal %>" 
                                                                        type="number" step="0.01" placeholder="0.00" 
                                                                        value="<%= t("PrecioOriginal") %>" 
                                                                        required
                                                                    >
                                                                </td>

                                                                <td>
                                                                    <input class="field frame" style="width: 100%; text-align: center;" 
                                                                        name="<%= nombreFecha %>" 
                                                                        id="<%= nombreFecha %>" 
                                                                        type="text" 
                                                                        value="<%= fechaFormulario(t("Fecha")) %>"
                                                                    >
                                                                </td>                                                                                                                            
                                                            <%                                       
                                                        Case 4
                                                            '
                                                            ' Si es MultiPrecios, aparecen dos precios distintos
                                                            '     

                                                            tot_Precio = tot_Precio + cDbl(t("PrecioOriginal"))
                                                            tot_Cambio = tot_Cambio + cDbl(t("Precio"))                                    

                                                            %>
                                                                <td>
                                                                    <input class="field frame" style="width:100%;" name="<%= nombreItem %>" id="<%= nombreItem %>" type="text" value="<%= t("Item") %>">
                                                                </td>

                                                                <td>
                                                                    <input class="field frame"
                                                                        style="width:100%; text-align: right;" 
                                                                        name="<%= nombrePrecioOriginal %>" 
                                                                        id="<%= nombrePrecioOriginal %>" class="gradeB_<%= tipoLinea %>"
                                                                        type="number" step="0.01" 
                                                                        placeholder="0.00" 
                                                                        value="<%= t("PrecioOriginal") %>" 
                                                                        onChange="CambiarMonedaLista('<%= t("LocalOrigen") %>', '<%= t("LocalDestino") %>', 1, '<%= nombrePrecioOriginal %>', '<%= nombrePrecio %>');">
                                                                </td>

                                                                <td>
                                                                    <input class="field frame"
                                                                        style="width:100%; text-align: right;" 
                                                                        name="<%= nombrePrecio %>" 
                                                                        id="<%= nombrePrecio %>"
                                                                        type="number" step="0.01" 
                                                                        placeholder="0.00" 
                                                                        value="<%= t("Precio") %>" 
                                                                        onChange="CambiarMonedaLista('<%= t("LocalDestino") %>', '<%= t("LocalOrigen") %>', 1, '<%= nombrePrecio %>', '<%= nombrePrecioOriginal %>');">
                                                                </td>

                                                                <td>
                                                                    <input class="field frame" style="width:100%; text-align: center;" name="<%= nombreFecha %>" id="<%= nombreFecha %>" type="text" value="<%= fechaFormulario(t("Fecha")) %>">
                                                                </td>         
                                                            <%                                             
                                                    End Select

                                                    response.write "<td style='text-align: center;'>"
                                                        %>
                                                            <button class="form-btn rojo" type="button" onclick="borrar('<%= t("Secuencia") %>', '<%= t("Item") %>')" >
                                                                <i class="fa fa-trash fa-xl" title='Borrar Item'></i>
                                                            </button>
                                                        <%
                                                    response.write "</td>"                     
                                                response.write "</tr>"

                                                t.MoveNext
                                            Loop Until t.eof
                                        end if
                                    %>
                                </tbody>  

                                <tfoot>
                                    <tr>
                                        <%
                                            if Cuenta = 1 then
                                                if MultiPrecio = 1 then
                                                    %>
                                                        <td class="sticky">
                                                            <%
                                                                if cuantos = 0 then
                                                                    response.write "Esta lista no tiene items"
                                                                else
                                                                    response.write "Se han encontrado " & cuantos & " items en la lista"
                                                                end if
                                                            %>                                                
                                                        </td>

                                                        <td class="sticky" style="text-align: right;"><%= FORMATNUMBER(tot_Precio, 2) %></td>
                                                        <td class="sticky" style="text-align: right;"><%= FORMATNUMBER(tot_Cambio, 2) %></td>
                                                        <td class="sticky">&nbsp;</td>
                                                    <%                
                                                else
                                                    %>
                                                        <td>
                                                            <%
                                                            if cuantos = 0 then
                                                                    response.write "Esta lista no tiene items"
                                                                else
                                                                    response.write "Se han encontrado " & cuantos & " items en la lista"
                                                                end if
                                                            %>                                                
                                                        </td>

                                                        <td class="sticky" style="text-align: right;"><%= FORMATNUMBER(tot_Precio, 2) %></td>
                                                        <td class="sticky">&nbsp;</td>                                                    
                                                    <%
                                                end if
                                            else
                                                if MultiPrecio = 1 then
                                                    %>
                                                        <td class="sticky">
                                                            <%
                                                                if cuantos = 0 then
                                                                    response.write "No se ha encontrado ningun item en la lista"
                                                                else
                                                                    response.write "Se han encontrado " & cuantos & " items en la lista"
                                                                end if
                                                            %>                                                
                                                        </td>

                                                        <td class="sticky" style="text-align: right;"><%= FORMATNUMBER(tot_Precio, 2) %></td>
                                                        <td class="sticky" style="text-align: right;"><%= FORMATNUMBER(tot_Cambio, 2) %></td>
                                                        <td colspan="2" class="sticky">&nbsp;</td>
                                                    <%     
                                                else
                                                    %>
                                                        <td class="sticky">
                                                            <%
                                                                if cuantos = 0 then
                                                                    response.write "No se ha encontrado ningun item en la lista"
                                                                else
                                                                    response.write "Se han encontrado " & cuantos & " items en la lista"
                                                                end if
                                                            %>                                                
                                                        </td>

                                                        <td class="sticky" style="text-align: right;"><%= FORMATNUMBER(tot_Precio, 2) %></td>
                                                        <td colspan="2" class="sticky">&nbsp;</td>                                                    
                                                    <%  
                                                end if
                                            end if
                                        %>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                    </div>
                </form>
            <!-- Fin de Lista de Items -->

            <!-- Formulario de Adición de Item -->
                <table class="subformulario" style="width: 100%; padding: 0px;">
                    <%
                        Select Case filaTipo
                            Case 1
                                %>
                                    <tr>
                                        <td class="encabezado">Nombre del Item</td>
                                        <td class="encabezado" style="width: 15%;">Precio Original</td>
                                        <td class="encabezado" style="width: 15%;">Cambio Monetario</td>
                                        <td class="encabezado" style="width:  5%;">&nbsp;</td>
                                    </tr>

                                    <tr>
                                        <td><input class="field" style="width: 100%;"                    id="n_Item"            name="n_Item"           type="text"                 placeholder="Item..."></td>
                                        <td><input class="field" style="width: 100%; text-align: right;" id="n_PrecioOriginal"  name="n_PrecioOriginal" type="number" step="0.01"   placeholder="0.00"  OnChange="CambiarMoneda('<%= MonedaOrigen %>','<%= MonedaDestino %>', 1);"></td>
                                        <td><input class="field" style="width: 100%; text-align: right;" id="n_Precio"          name="n_Precio"         type="number" step="0.01"   placeholder="0.00"  OnChange="CambiarMoneda('<%= MonedaDestino %>','<%= MonedaOrigen %>', 2);"></td>
                                        <td style="text-align: center;" >
                                            <button class="form-btn verde" type="button" onclick="grabarNuevoItem()">
                                                <i class="fa fa-save"></i>
                                            </button>
                                        </td>
                                    </tr>                                
                                <%
                            Case 2
                                %>
                                    <tr>
                                        <td class="encabezado">Nombre del Item</td>
                                        <td class="encabezado" style="width: 15%;">Precio Original</td>
                                        <td class="encabezado" style="width:  5%;">&nbsp;</td>
                                    </tr>                                    

                                    <tr>
                                        <td><input class="field" style="width: 100%;"                     id="n_Item"            name="n_Item"           type="text"                 placeholder="Item..."></td>
                                        <td><input class="field" style="width: 100%; text-align: right;"  id="n_PrecioOriginal"  name="n_PrecioOriginal" type="number" step="0.01"   placeholder="0.00"  OnChange="CambiarMoneda('<%= MonedaOrigen %>','<%= MonedaDestino %>', 1);"></td>

                                        <td style="text-align: center;" >
                                            <button class="form-btn verde" type="button" onclick="grabarNuevoItem()">
                                                <i class="fa fa-save"></i>
                                            </button>
                                        </td>
                                    </tr>                                                                         
                                <%                            
                            Case 3
                                %>
                                    <tr>
                                        <td class="encabezado">Nombre del Item</td>
                                        <td class="encabezado" style="width: 15%;">Precio Original</td>                                        
                                        <td class="encabezado" style="width: 15%;">Cambio Monetario</td>
                                        <td class="encabezado" style="width: 15%;">Fecha</td>
                                        <td class="encabezado" style="width:  5%;">&nbsp;</td>
                                    </tr>                                    

                                    <tr>
                                        <td><input class="field" style="width: 100%;"                     id="n_Item"            name="n_Item"           type="text"                 placeholder="Item..."></td>
                                        <td><input class="field" style="width: 100%; text-align: right;"  id="n_PrecioOriginal"  name="n_PrecioOriginal" type="number" step="0.01"   placeholder="0.00"  OnChange="CambiarMoneda('<%= MonedaOrigen %>','<%= MonedaDestino %>', 1);"></td>
                                        <td><input class="field" style="width: 100%; text-align: right;"  id="n_Precio"          name="n_Precio"         type="number" step="0.01"   placeholder="0.00"  OnChange="CambiarMoneda('<%= MonedaDestino %>','<%= MonedaOrigen %>', 2);"></td>
                                        <td><input class="field" style="width: 100%; text-align: center;" id="n_Fecha"           name="n_Fecha"          type="text"                 placeholder="dd/mm/aaaa" ></td>

                                        <td style="text-align: center;" >
                                            <button class="form-btn verde" type="button" onclick="grabarNuevoItem()">
                                                <i class="fa fa-save"></i>
                                            </button>
                                        </td>
                                    </tr>                         
                                <%                               
                            Case 4
                                %>
                                    <tr>
                                        <td class="encabezado">Nombre del Item</td>
                                        <td class="encabezado" style="width: 15%;">Precio Original</td>                                        
                                        <td class="encabezado" style="width: 15%;">Fecha</td>
                                        <td class="encabezado" style="width:  5%;">&nbsp;</td>
                                    </tr>  

                                    <tr>
                                        <td><input class="field" style="width: 100%;"                     id="n_Item"            name="n_Item"           type="text"                 placeholder="Item..."></td>
                                        <td><input class="field" style="width: 100%; text-align: right;"  id="n_PrecioOriginal"  name="n_PrecioOriginal" type="number" step="0.01"   placeholder="0.00"  OnChange="CambiarMoneda('<%= MonedaOrigen %>','<%= MonedaDestino %>', 1);"></td>
                                        <td><input class="field" style="width: 100%; text-align: center;" id="n_Fecha"           name="n_Fecha"          type="text"                 placeholder="dd/mm/aaaa" ></td>

                                        <td style="text-align: center;" >
                                            <button class="form-btn verde" type="button" onclick="grabarNuevoItem()">
                                                <i class="fa fa-save"></i>
                                            </button>
                                        </td>
                                    </tr>
                                <%                               
                        End Select
                    %>
                </table>
            <!-- Fin de Formulario de Adición de Item -->
        </div>

        <br /><br />           

        <script>
            function grabarNuevoItem(){
                var item, precio, cambio, fecha, vinculo;

                var codigo = "<%= Codigo %>";
                var multi = <%= MultiPrecio %>;
                var cuenta = <%= Cuenta %>;

                if (cuenta == 1) {
                    if (multi == 1 ){
                        var item = document.getElementById("n_Item").value;
                        var precio = document.getElementById("n_PrecioOriginal").value;
                        var cambio = document.getElementById("n_Precio").value;
                    }
                    else {
                        var item = document.getElementById("n_Item").value;
                        var precio = document.getElementById("n_PrecioOriginal").value;
                    };
                }
                else {
                    if (multi == 1 ){
                        var item = document.getElementById("n_Item").value;
                        var precio = document.getElementById("n_PrecioOriginal").value;
                        var cambio = document.getElementById("n_Precio").value;
                        var fecha = document.getElementById("n_Fecha").value;
                    }
                    else {
                        var item = document.getElementById("n_Item").value;
                        var precio = document.getElementById("n_PrecioOriginal").value;
                        var fecha = document.getElementById("n_Fecha").value;
                    };          
                };

                vinculo = "listas_items_grabar.asp?cod=" + codigo + "&m=" + multi + "&q=" + cuenta + "&i="  + item + "&p1=" + precio + "&p2=" + cambio + "&f=" + fecha               
                window.location.href = vinculo;
            }

            function borrar(secuencia, nombre) {
                var confirmacion = confirm("Está seguro de borrar el item '" + nombre + "' de la lista?");
                var vinculo = "listas_items_delete.asp?s=" + secuencia;

                if (confirmacion) {
                    window.location.href = vinculo;
                } else {
                    alert("Proceso Cancelado");
                }        
            };

            function abrir(vinculo) {
                window.location.href = vinculo;                
            }

            function CambiarMoneda(desde, hasta, m) {
                CambiarMonedaLista(desde, hasta, m, "n_PrecioOriginal", "n_Precio");          
            };

            function CambiarMonedaLista(desde, hasta, m, nomOriginal, nomPrecio) {
                <%
                    dim cc2, tloc2, tsim2, tfor2

                    set cc2 = server.CreateObject("ADODB.Connection")
                    cc2.open Application("Conn")
                %>     

                var donde = 0;
                var k = 0;
                var PrecioOriginal = 0.00;
                var PrecioDestino = 0.00;

                if (m == 1) {
                    PrecioOriginal = document.getElementById(nomOriginal).value;

                    if (PrecioOriginal < 0) {
                        PrecioOriginal = (-1 * PrecioOriginal);
                        alert("El Valor SIEMPRE debe ser mayor o igual a cero.");
                        document.getElementById(nomOriginal).value = PrecioOriginal;
                    };          
                };

                if (m == 2) {
                    PrecioOriginal = document.getElementById(nomPrecio).value

                    if (PrecioOriginal < 0) {
                        PrecioOriginal = (-1 * PrecioOriginal);
                        alert("El Valor SIEMPRE debe ser mayor o igual a cero.");
                        document.getElementById(nomPrecio).value = PrecioOriginal;
                    };          
                };

                var formatter = new Intl.NumberFormat('en-US', {
                    style: 'decimal',
                    currency: 'USD',
                });   

                var locales = [<%
                    set tloc2 = cc2.execute("SELECT Local, Simbolo, Formula " & _
                                            "FROM seg_Cripto_NumParse_Locales " & _
                                        "WHERE Local <> 'NUM' " & _
                                        "ORDER BY Local ASC;")

                    if not (tloc2.bof or tloc2.eof) then
                        response.write "'*'"

                        do
                            response.write ", "
                            response.write "'" & tloc2("local") & "'"
                            tloc2.MoveNext
                        loop until (tloc2.eof)
                    end if

                    tloc2.close: set tloc2 = nothing
                %>];

                var simbolos = [<%
                    set tsim2 = cc2.execute("SELECT Local, Simbolo, Formula " & _
                                            "FROM seg_Cripto_NumParse_Locales " & _
                                        "WHERE Local <> 'NUM' " & _
                                        "ORDER BY Local ASC;")

                    if not (tsim2.bof or tsim2.eof) then
                        response.write "'*'"

                        do
                            response.write ", "
                            response.write "'" & tsim2("simbolo") & "'"
                            tsim2.MoveNext
                        loop until (tsim2.eof)
                    end if

                    tsim2.close: set tsim2 = nothing
                %>];

                var formula = [<%
                    set tfor2 = cc2.execute("SELECT Local, Simbolo, Formula " & _
                                            "FROM seg_Cripto_NumParse_Locales " & _
                                        "WHERE Local <> 'NUM' " & _
                                        "ORDER BY Local ASC;")

                    if not (tfor2.bof or tfor2.eof) then
                        response.write "'*'"

                        do
                            response.write ", "
                            response.write "'" & tfor2("formula") & "'"
                            tfor2.MoveNext
                        loop until (tfor2.eof)
                    end if

                    tfor2.close: set tfor2 = nothing
                %>];
                
                if (desde == hasta) {
                    if (m == 1) {document.getElementById(nomPrecio).value = PrecioOriginal};
                    if (m == 2) {document.getElementById(nomOriginal).value = PrecioOriginal};
                } 
                else {
                    /*
                        Llevamos el PrecioOriginal a USD
                    */

                    donde = 0;

                    for(let k = 0; k < locales.length; k++) {
                        if (locales[k] == desde) {
                            donde = k
                        }
                    };

                    PrecioDestino = (PrecioOriginal / formula[donde]);

                    /*
                        Llevamos el PrecioOriginal USD a Moneda Destino
                    */

                    donde = 0;

                    for(let k = 0; k < locales.length; k++) {
                        if (locales[k] == hasta) {
                            donde = k
                        }
                    };

                    PrecioDestino = (PrecioDestino * formula[donde]);
                    /* PrecioDestino = formatter.format(PrecioDestino); */
                    PrecioDestino = PrecioDestino.toFixed(2) 

                    if (m == 1) { document.getElementById(nomPrecio).value = PrecioDestino };
                    if (m == 2) { document.getElementById(nomOriginal).value = PrecioDestino };
                }
                <%
                    cc2.close: set cc2 = nothing 
                %>              
            };

            function enviarFormulario(){
                document.getElementById("form_lista").submit(); 
            }  
            
            mask(document.getElementById('n_Fecha'), ['99/99/9999']);  

            <%
                ssql = "SELECT RIGHT('00000000000000000000' + CAST(d.Secuencia AS varchar(18)), 18) AS Llave " & _
                            "FROM pre_Listas_Detalles AS d " & _
                        "WHERE (d.Usuario = '" & Request.Cookies("Usuario") & "') " & _
                            "AND (d.Codigo = '" & Codigo & "') " & _
                        "ORDER BY d.Item;"

                set f = conn.execute(ssql)
                    if not (f.bof or f.eof) then
                        cuantos = 0   

                        do
                            cuantos = cuantos + 1
                            nombreFecha = "Lfech_" & f("Llave")

                            response.write "mask(document.getElementById('" & nombreFecha & "'), ['99/99/9999']); "
                            
                            f.MoveNext
                        loop until f.eof
                    end if       
                    
                f.close: set f = nothing
            %>         
        </script>

        <% conn.close: set conn = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->        
    </body>
</html>