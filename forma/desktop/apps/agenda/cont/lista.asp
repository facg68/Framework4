<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->
        <title>Lista de Contactos</title>
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0050"
            SysLockOut

            function pOrdenar(CampoComparacion, CampoOrdenamiento, DirOrdenamiento)
                if CampoComparacion = CampoOrdenamiento then
                    if DirOrdenamiento = "ASC" then
                        pOrdenar = "DESC"
                    else
                        pOrdenar = "ASC"
                    end if
                else
                    pOrdenar = "ASC"
                end if
            end function
        %>   

        <style>
            td.reset {
                font-family: Ruda;
                font-size: 16px;
                line-height: 20px;

                margin: 0;
                padding: 0;

                background: transparent;

                letter-spacing: normal;
                word-spacing: normal;
                white-space: normal;
            }
        </style>    
    </head>

    <body plantilla="tabla" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            dim con, t, sqlString, vinculo, usu
            dim tipo, categ, orden, dir, apendice

            usu = Request.Cookies("Usuario")

            '
            ' Primero analizamos el objeto FORM
            '
            tipo = Request.Form("cboTipo")
            categ = Request.Form("cboCategoria")
            orden = Request.Form("orden")
            dir = Request.Form("dir")
            visibilidad = Request.Form("cboVisibilidad")    

            if tipo = "" then 
            '
            ' Analizamos el Objeto QueryString...
            '
            tipo = Request.QueryString("t")
            categ = Request.QueryString("c")
            orden = Request.QueryString("o1")
            dir = Request.QueryString("o2")
            visibilidad = Request.QueryString("v")
            end if

            if tipo = "" then tipo = "PE"
            if categ = "" then categ = "principal"
            if orden = "" then orden = "nombre"
            if dir = "" then dir = "ASC"
            if visibilidad = "" then visibilidad = 1

            apendice = "&t=" & tipo & "&c=" & categ & "&o1=" & orden & "&o2=" & dir & "&v=" & visibilidad

            sqlString = "SELECT DISTINCT Usuario, Codigo, Nombre, Correo, Cumple, Telefono " & _
                        "FROM ( " & _
                                " SELECT Usuario, Codigo, Nombre, Correo, Cumple, Telefono, TipoContacto, Categ " & _
                                " FROM con_FiltroContactos " & _
                                " WHERE (Codigo <> '" & usu & "') " & _
                                " AND (Usuario = '" & usu & "') " 

            if visibilidad <> "*" then sqlString = sqlString & " AND (Visible = '" & visibilidad & "') " 
            if tipo <> "*" then sqlString = sqlString & " AND (TipoContacto = '" & tipo & "') " 
            if categ <> "*" then sqlString = sqlString & " AND (Categ = '" & categ & "') " 

            sqlString = sqlString & ") as t ORDER BY " & orden & " " & dir & ";"

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")
            set t = con.Execute(sqlString)        
        %>        

        <br />

        <form id="formulario" name="formulario" method="post" action="lista.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 20%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Contactos
                </div>
                
                <div style="flex: 0 0 80%; text-align: right;">
                    <select class="no-field" name="cboVisibilidad" id="cboVisibilidad" onchange="Requery()">
                        <option value="1" <% if visibilidad = "1" then response.write " selected" %>>Activos</option>
                        <option value="0" <% if visibilidad = "0" then response.write " selected" %>>Obsoletos</option>
                        <option value="*" <% if visibilidad = "*" then response.write " selected" %>>Completa</option>
                    </select>

                    &nbsp;

                    <select class="no-field" name="cboTipo" id="cboTipo" onchange="RequeryCategs()">
                        <option value="*">- - Todos - -</option>
                        <%
                            sqlString ="SELECT Codigo, Nombre " & _
                                        "FROM con_Contactos_Tipos " & _
                                        "WHERE (Usuario = '" & usu & "') " & _
                                    "ORDER BY Nombre ASC;"

                            set cbox = con.execute(sqlString)

                            if not (cbox.bof or cbox.eof) then
                                Do
                                    response.write "<option value='" & cbox("Codigo") & "' "
                                        if Tipo = cbox("Codigo") then 
                                            response.write " selected='selected'" 
                                        end if
                                    response.write ">" & cbox("Nombre") & "</option>"

                                    cbox.MoveNext
                                Loop Until cbox.eof
                            end if

                            cbox.close: set cbox = nothing
                        %>
                    </select>  

                    &nbsp;

                    <select class="no-field" name="cboCategoria" id="cboCategoria" onchange="Requery()">
                        <option value="*">- - Todos - -</option>
                        <%
                            sqlString = "SELECT Codigo, Nombre " & _
                                        "FROM con_Contactos_Categorias " & _
                                        "WHERE (Usuario = '" & usu & "') " & _
                                        "AND (Tipo = '" & Tipo & "') " & _
                                    "ORDER BY Nombre ASC;"

                            set cbox = con.execute(sqlString)

                            if not (cbox.bof or cbox.eof) then
                                Do
                                    response.write "<option value='" & cbox("Codigo") & "' "
                                        if categ = cbox("Codigo") then 
                                            response.write " selected='selected'" 
                                        end if
                                    response.write ">" & cbox("Nombre") & "</option>"

                                    cbox.MoveNext
                                Loop Until cbox.eof
                            end if

                            cbox.close: set cbox = nothing
                        %>
                    </select>     

                    &nbsp;

                    <button class='form-btn verde normal' type='button' onclick="NuevoContacto()">Nuevo</button>
                </div>
            </div>        

            <div class="main" style="width: 95%;">
                <div class="no-ver">
                    <input id="campoOrdenamiento"       name="campoOrdenamiento"        type="text"     value="<%= orden %>">
                    <input id="direccionOrdenamiento"   name="direccionOrdenamiento"    type="text"     value="<%= dir %>">  
                    <input id="ordenamiento"            name="ordenamiento"             type="text"     value="<%= Orden %>">
                    <input id="direccion"               name="direccion"                type="text"     value="<%= dir %>">                                  
                </div>

                <div class="line">
                    <div class="tabla-wrapper">
                        <table class="tabla tabla-violet">
                            <thead>
                                <tr>
                                    <th class="sticky" style="width:  5%; text-align: center;">&nbsp;</th>
                                    <th class="sticky" style="width: 30%; text-align: center;" onClick="ordenar('nombre',   '<%= pOrdenar("nombre", orden, dir)   %>');">Nombre</th>
                                    <th class="sticky" style="width: 25%; text-align: center;" onClick="ordenar('correo',   '<%= pOrdenar("correo", orden, dir)   %>');">Correo</th>
                                    <th class="sticky" style="width: 10%; text-align: center;" onClick="ordenar('cumple',   '<%= pOrdenar("cumple", orden, dir)   %>');">Cumpleaños</th>
                                    <th class="sticky" style="width: 15%; text-align: center;" onClick="ordenar('telefono', '<%= pOrdenar("telefono", orden, dir) %>');">Telefono</th>
                                    <th class="sticky" style="width: 15%; text-align: center;">&nbsp;</th>
                                </tr>
                            </thead>

                            <tbody>
                                <%
                                    if not (t.bof or t.eof) then
                                        Do
                                            cuantos = cuantos + 1
                                            vinculo = "cont_editar.asp?con=" & t("Codigo") & apendice

                                            %>
                                                <tr>
                                                    <td class="reset" style="text-align: center;" onclick="irA('<%= vinculo %>')">
                                                        <img src="<%= request.Cookies("usuPath") & "/fotos/" & t("Codigo") & "_s.jpg" %>" 
                                                            onerror="this.src='/core/imagenes/misc/foto.jpg'" width="50px">
                                                    </td>

                                                    <td class="reset" style="text-align: left; padding: 8px;" onclick="irA('<%= vinculo %>')">
                                                        <%= t("Nombre") %>
                                                    </td>

                                                    <td class="reset" style="text-align: center;" onclick="irA('<%= vinculo %>')">
                                                        <%= t("Correo") %>
                                                    </td>

                                                    <td class="reset" style="text-align: center;" onclick="irA('<%= vinculo %>')">
                                                        <%= t("Cumple") %>
                                                    </td>

                                                    <td class="reset" style="text-align: center;" onclick="irA('<%= vinculo %>')">
                                                        <%= t("Telefono") %>
                                                    </td>

                                                    <td class="reset" style="text-align: center;">
                                                        <button type="button" class="form-btn rojo" onclick="borrar('<%= t("Codigo") %>', '<%= t("Nombre") %>')">
                                                            <i class='fa fa-trash fa-xl'></i>
                                                        </button>

                                                        <button type="button" class="form-btn azul" onclick="obsoleto('<%= t("Codigo") %>', '<%= t("Nombre") %>')">
                                                            <i class='fa fa-random fa-xl'></i>
                                                        </button>
                                                    </td>
                                                </tr>
                                            <%

                                            t.MoveNext
                                        Loop Until t.eof
                                    end if

                                    t.close: set t = nothing
                                %>
                            </tbody>

                            <tfoot>
                                <tr>
                                    <td class="sticky" style="text-align: center;" colspan="6">
                                        <%
                                            Select Case cuantos
                                                case 0: response.write "No se encontraron contactos"
                                                case 1: response.write "S&oacute;lo se encontr&oacute; un Comtacto"
                                                case else
                                                    response.write "Se encontraron " & Cuantos &  " Contactos"                                
                                            end Select
                                        %>                                    
                                    </td>
                                </tr>
                            </tfoot>
                        </table>
                    </div>
                </div>
            </div>
        </form>
  
        <br /><br />   

        <script type="text/javascript">
            function RequeryCategs() {
                document.getElementById("cboCategoria").value = "*";
                Requery();
            }

            function Requery() {
                document.getElementById("formulario").submit();
            }

            function ordenar(campo, direccion) {
                var tipo = document.getElementById("cboTipo").value;
                var categ = document.getElementById("cboCategoria").value;
                var vis = document.getElementById("cboVisibilidad").value;
                var vinculo = "";

                vinculo = "lista.asp?v=" + vis + "&t=" + tipo + "&c=" + categ + "&o1=" + campo + "&o2=" + direccion;
                window.location.href = vinculo;
            }

            function obsoleto(codigo, nombre) {
                var tipo = document.getElementById("cboTipo").value;
                var categ = document.getElementById("cboCategoria").value;
                var vis = document.getElementById("cboVisibilidad").value;
                var campo = document.getElementById("campoOrdenamiento").value;
                var direccion = document.getElementById("direccionOrdenamiento").value;
                var vinculo = "";

                var confirmacion = confirm("Esta seguro de cambiar el estatus del contacto '" + nombre + "'?");
                vinculo = "cont_obsoleto.asp?con=" + codigo + "&v=" + vis + "&t=" + tipo + "&c=" + categ + "&o1=" + campo + "&o2=" + direccion;        

                if (confirmacion) {
                    window.location.href = vinculo;
                } else {
                    window.alert("Proceso Cancelado.");        
                }        
            }      

            function NuevoContacto() {
                window.location.href = "cont_nuevo.asp";
            }

            function irA(direccion) {
                window.location.href = direccion;
            } 
            
            function borrar(codigo, nombre) {
                var tipo = document.getElementById("cboTipo").value;
                var categ = document.getElementById("cboCategoria").value;
                var vis = document.getElementById("cboVisibilidad").value;
                var campo = document.getElementById("campoOrdenamiento").value;
                var direccion = document.getElementById("direccionOrdenamiento").value;
                var vinculo = "";

                var confirmacion = confirm("Esta seguro de borrar el contacto '" + nombre + "'?");
                vinculo = "cont_borrar.asp?con=" + codigo + "&v=" + vis + "&t=" + tipo + "&c=" + categ + "&o1=" + campo + "&o2=" + direccion;        

                if (confirmacion) {
                    window.location.href = vinculo;
                } else {
                    window.alert("Proceso Cancelado.");        
                }        
            }            
        </script> 

        <% con.close: set con = nothing  %> 
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>