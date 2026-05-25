<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <title>Editar Tiendas</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        
        <%
            thisSystem = "discos"
            thisProcess = "discos.0210"
            SysLockOut

            dim con, t, tt, sqlString, data, labels
            dim cbox, cuantos, ordenamiento, oo, vv, Grupo
            dim Codigo, Nombre, Descripcion, Cuenta, vinculo, verTipo   

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")     
        %>         

        <style>
            body { overflow-y: hidden; }

            a.linea, a.linea:link, a.linea:visited,
            a.linea:focus, a.linea:hover, 
            a.linea:active { color: black; }
        </style>   
    </head>

    <body plantilla="dividida" tabla="55" grafica="45" reserva="200">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            Usuario = Request.Cookies("usuario")
            Grupo = Request.QueryString("g")
        
            cuantos = 0
            if Grupo = "" then Grupo = "1"

            sqlString = "SELECT t.Usuario, t.Codigo, t.Nombre, t.Contacto, t.SitioWeb, t.Telefono1, p.Nombre AS Pais, " & _
                              " CASE WHEN t .Tipo = 0 THEN 'Online' WHEN t .Tipo = 1 THEN 'Fisica' END AS TipoTienda, " & _
                              " ISNULL(m.Cuantos, 0) AS Medios, t.Estatus " & _
                         "FROM dbo.discos_Tiendas AS t " & _
                   "INNER JOIN dbo.seg_Paises AS p " & _
                           "ON t.Pais = p.Codigo " & _
              "LEFT OUTER JOIN (" & _
                                    "SELECT Usuario, Tienda, COUNT(*) AS Cuantos " & _
                                      "FROM dbo.discos_Paquetes " & _
                                  "GROUP BY Usuario, Tienda) AS m " & _
                           "ON t.Usuario = m.Usuario " & _
                          "AND t.Codigo = m.Tienda " & _
                        "WHERE (t.Codigo <> '00000000') " & _
                          "AND (t.Usuario = '" & Request.Cookies("Usuario") & "') " 

            if Grupo <> "*" then
                sqlString = sqlString & "AND (t.Estatus = '" & Grupo & "') "
            end if

            sqlString =sqlString & "ORDER BY t.Nombre ASC;"

            set t = con.execute(sqlString)
        %>          

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                Tiendas
            </div>
            
            <div style="flex: 0 0 50%; text-align: right;">
                <select class="no-field" name="cboVerGrupo" id="cboVerGrupo" onChange="requery();">
                    <option value="1" <% if Grupo = "1" then response.write " selected" %>>Tiendas Activas</option>
                    <option value="0" <% if Grupo = "0" then response.write " selected" %>>Tiendas Cerradas</option>
                    <option value="*" <% if Grupo = "*" then response.write " selected" %>>Todas las Tiendas</option>                    
                </select>                   

                <button class="form-btn verde" onclick="nuevo()">
                    <i class='fa fa-edit fa-xl' title='Nuevo'></i>
                </button>                   
            </div>
        </div>        

        <div class="main">
            <div class="line">
                <div class="tabla-wrapper">
                    <table class="tabla tabla-blue">
                        <thead>
                            <tr>
                                <th class="sticky" style="width: 18%; text-align: center;">Nombre</th>
                                <th class="sticky" style="width: 18%; text-align: center;">Contacto</th>
                                <th class="sticky" style="width: 14%; text-align: center;">Teléfono</th>
                                <th class="sticky" style="width: 18%; text-align: center;">País</th>
                                <th class="sticky" style="width:  6%; text-align: center;">Tipo</th>
                                <th class="sticky" style="width:  6%; text-align: center;">Medios</th>
                                <th class="sticky" style="width: 20%; text-align: center;">&nbsp</th>
                            </tr>
                        </thead>

                        <tbody>
                            <%
                                if not (t.bof or t.eof) then
                                    sw = -1

                                    Do
                                        sw = -1 * sw 
                                        cuantos = cuantos + 1
                                        vinculo = "editar.asp?c=" & t("Codigo") & "&g=" & Grupo
                                        classLinea = ""

                                        if t("Medios") > 0 then
                                            classLinea = "disabled"
                                        end if   
                                            %>
                                                <tr>
                                                    <td style="text-align: left;">
                                                        <a class="linea" href="<%= vinculo %>">
                                                            <%= t("Nombre") %>
                                                        </a>
                                                    </td>

                                                    <td style="text-align: left;">
                                                        <a class="linea" href="<%= vinculo %>">
                                                            <%= t("Contacto") %>
                                                        </a>
                                                    </td>

                                                    <td style="text-align: center;">
                                                        <a class="linea" href="<%= vinculo %>">
                                                            <%= t("Telefono1") %>
                                                        </a>
                                                    </td>

                                                    <td style="text-align: left;">
                                                        <a class="linea" href="<%= vinculo %>">
                                                            <%= t("Pais") %>
                                                        </a>
                                                    </td>

                                                    <td style="text-align: left;">
                                                        <a class="linea" href="<%= vinculo %>">
                                                            <%= t("TipoTienda") %>
                                                        </a>
                                                    </td>

                                                    <td style="text-align: center;">
                                                        <a class="linea" href="<%= vinculo %>">
                                                            <%= t("Medios") %>
                                                        </a>
                                                    </td>

                                                    <td style="text-align:center; padding: 2px;">
                                                        <a href="../../lista_paquetes.asp?f=<%= "Tienda = '" & t("Codigo") & "' " %>&t=<%= "Tienda = '" & t("Nombre") & "'"  %>" >
                                                            <button class="form-btn azul">
                                                                <i class="fa fa-eye fa-xl" title="Filtrar"></i>
                                                            </button>
                                                        </a>
                                               
                                                        <% if classLinea = "disabled" then %>
                                                            <button class="form-btn rojo disabled">
                                                                <i class="fa fa-trash fa-xl" title="Borrar"></i>
                                                            </button>
                                                        <% else %>
                                                            <a onclick="borrar('<%= t("Codigo") %>')">
                                                                <button class="form-btn rojo">
                                                                    <i class=" fa fa-trash fa-xl" title="Borrar"></i>
                                                                </button>
                                                            </a>
                                                        <% end if %> 
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
                                <td class="sticky" colspan="7" style="text-align: center;">
                                    <%
                                        if cuantos = 0 then
                                            response.write "No se ha encontrado ninguna tienda."
                                        else
                                            response.write "Se han encontrado " & cuantos & " tiendas"
                                        end if
                                    %>                                  
                                </td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>

            <%
                sqlString = "exec discos_gfx_listaTiendas '" & Usuario & "'"
                apexColumns "", "chart", sqlString, "Nombre", "Cuantos", "#3c7fb6ff", 150                 
            %>
        </div>
  
        <br /><br />  

        <script>
            function requery() {
                var grupo = document.getElementById("cboVerGrupo").value;
                var vinculo = "lista.asp?g=" + grupo;

                window.location.href = vinculo;
            }

            function nuevo() {
                var grupo = document.getElementById("cboVerGrupo").value;
                var vinculo = "nuevo.asp?g=" + grupo;

                window.location.href = vinculo;
            }            

            function borrar(codigo) {
                var confirmacion = confirm("Está seguro de borrar el registro seleccionado?");

                var grupo = document.getElementById("cboVerGrupo").value;
                var vinculo = "borrar.asp?c=" + codigo + "&g=" + grupo;

                if (confirmacion) {
                    window.location.href = vinculo;
                } else {
                    alert("Proceso Cancelado.");        
                }        
            }                  
        </script> 

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>    
</html>