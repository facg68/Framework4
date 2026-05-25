<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <title>Editar Listas</title>    
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->             
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0110"
            SysLockOut


            function EstaVacia(Lista)
                dim cc, tt, ssql

                ssql = "SELECT COUNT(*) AS Cuantos " & _
                       "FROM pre_Listas_Detalles " & _
                       "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                       "AND (Codigo = '" & Lista & "');"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                    set tt = cc.execute(ssql)
                        EstaVacia = 0

                        if tt("Cuantos") = 0 then
                            EstaVacia = 1
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

            .col1 {
                white-space: nowrap;
                font-weight: bold;
                min-width: fit-content;
            }

            .col2 { flex: 0 0 10%; }
            .col3 { flex: 0 0 15%; }
            .col4 { flex: 1; }
            .col5 { flex: 0 0 60px; }

            a.linea, a.linea:link, a.linea:visited,
            a.linea:focus, a.linea:hover, 
            a.linea:active { color: black; }                
        </style>           
    </head>

    <body plantilla="tabla" reserva="210">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            dim con, t, sqlString, cbox, cuantos, Grupo
            dim Codigo, Nombre, Descripcion, Cuenta

            Usuario = Request.Cookies("usuario")
            Grupo = Request.QueryString("g")
        
            cuantos = 0
            if Grupo = "" then Grupo = "A"

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            sqlString = "SELECT Codigo, Nombre, Descripcion, Cuenta " & _
                        "FROM pre_Listas_Encabezado " & _
                        "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') "

            if Grupo <> "*" then
                sqlString =sqlString & "AND (Grupo = '" & Grupo & "') "
            end if

            sqlString =sqlString & "ORDER BY Nombre ASC;"

            set t = con.execute(sqlString)
        %>       

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                Listas
            </div>
            
            <div style="flex: 0 0 50%; text-align: right;">
                <select class="field normal" name="cboVerGrupo" id="cboVerGrupo" onChange="Requery();">
                    <option value="A" <% if Grupo = "A" then response.write " selected" %>>Listas Activas</option>
                    <option value="W" <% if Grupo = "W" then response.write " selected" %>>Listas En Espera</option>
                    <option value="S" <% if Grupo = "S" then response.write " selected" %>>Listas Archivadas</option>
                    <option value="*" <% if Grupo = "*" then response.write " selected" %>>Todas las Listas</option>
                </select> 
            </div>
        </div>        

        <div class="main" style="width: 95%;">
            <div class="line">
                <div class="tabla-wrapper">
                    <table class="tabla tabla-green">
                        <thead>
                            <tr>
                                <th class="sticky" style="width: 10%; text-align: center;">Codigo</td>
                                <th class="sticky" style="width: 30%; text-align: center;">Nombre</td>
                                <th class="sticky" style="width: 35%; text-align: center;">Descripcion</td>
                                <th class="sticky" style="width: 25%; text-align: center;">Acciones</td>                 
                            </tr>
                        </thead>

                        <tbody>                    
                            <%
                                if not (t.bof or t.eof) then
                                    Do
                                        cuantos = cuantos + 1

                                        response.write "<tr "
                                            if EstaVacia(t("Codigo")) = 1 then
                                                response.write " class='tr-rojo'"
                                            end if
                                        response.write ">"
                                            response.write "<td style='text-align: center;'>"
                                                response.write t("Codigo")
                                            response.write "</td>"                              

                                            response.write "<td>"
                                                response.write t("Nombre")
                                            response.write "</td>"                              

                                            response.write "<td style='text-align: left;'>"
                                                response.write t("Descripcion")
                                            response.write "</td>"                              

                                            response.write "<td style='text-align: center;'>"
                                                %>
                                                    <button class="form-btn verde" onclick="editar('<%= t("Codigo") %>')">
                                                        <i class="fa fa-edit fa-xl" title='Items en la lista'></i>
                                                    </button>

                                                    <button class="form-btn azul" onclick="items('<%= t("Codigo") %>')">
                                                        <i class="fa fa-list fa-xl" title='Items en la lista'></i>
                                                    </button>

                                                    <button class="form-btn rojo" onclick="borrar('<%= t("Codigo") %>')">
                                                        <i class="fa fa-trash fa-xl" title='Borrar lista'></i>
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
                                <td class="sticky" style="text-align: center;" colspan="4">
                                    <%
                                        if cuantos = 0 then
                                            response.write "No se ha encontrado ninguna lista."
                                        else
                                            response.write "Se han encontrado " & cuantos & " listas."
                                        end if
                                    %>
                                </td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>

            <form name="form_transaccion" id="form_transaccion" method="post" action="listas_grabar.asp">
                <div class="fila">
                    <div class="col1">Nueva Lista</div>

                    <div class="col2">
                        <input class="field" style="width: 100%; text-align: center;" id="nuevoCodigo" name="nuevoCodigo" type="text" placeholder="Código"/>
                    </div>

                    <div class="col3">
                        <input class="field" style="width: 100%; text-align: left;" id="nuevoNombre" name="nuevoNombre" type="text" placeholder="Nombre de Lista"/>
                    </div>

                    <div class="col4">
                        <input class="field" style="width: 100%; text-align: center;" id="nuevaDescripcion" name="nuevaDescripcion" type="text" placeholder="Descripción"/>
                    </div>

                    <div class="col5">
                        <button class="form-btn verde " type="submit">
                            <i class="fa fa-save fa-xl" title="Añadir"></i>
                        </button>   
                    </div>
                </div>             
            </form>
        </div>
  
        <br /><br />   

        <script type="text/javascript">
            function Requery() {
                var grupo = document.getElementById("cboVerGrupo").value;
                var vinculo = "lista.asp?g=" + grupo;

                window.location.href = vinculo;
            }

            function borrar(codigo) {
                var confirmacion = confirm("Está seguro de borrar la lista " + codigo + "?");
                var vinculo = "listas_borrar.asp?l=" + codigo;

                if (confirmacion) {
                    window.location.href = vinculo;
                } else {
                    alert("Proceso Cancelado.");        
                }        
            }    

            function editar(codigo) {
                var vinculo = "listas_editar.asp?l=" + codigo;
                window.location.href = vinculo;
            }

            function items(codigo) {
                var vinculo = "listas_items.asp?l=" + codigo;
                window.location.href = vinculo;
            }
        </script>

        <% con.close: set con = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>