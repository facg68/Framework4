<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Editar Casas</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "discos"
            thisProcess = "discos.0220"
            SysLockOut

            dim con, t, tt, sqlString, data, labels
            dim cbox, cuantos, ordenamiento, oo, vv
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

    <body plantilla="dividida" tabla="55" grafica="45" reserva="175">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            Usuario = Request.Cookies("usuario")
            verTipo = Request.QueryString("v")
            ordenamiento = Request.QueryString("o")
        
            cuantos = 0
            if ordenamiento = "" then ordenamiento = "1"
            if verTipo = "" then verTipo = "1"

            select case ordenamiento
                case 1: oo = "Nombre ASC;"
                case 2: oo = "Medios ASC;"
                case 3: oo = "Nombre DESC;"
                case 4: oo = "Medios DESC;"
            end select

            select case verTipo
                case 1: vv = "AND (c.Obsoleta = 0) "
                case 0: vv = "AND (c.Obsoleta = 1) "
            end select               

            sqlString = "SELECT c.Usuario, c.Codigo, c.Nombre, c.Musica, c.Video, c.Juegos, c.Software, c.Libros, c.Obsoleta, ISNULL(t.Cuantos, 0) AS Medios " & _
                          "FROM dbo.discos_Casas AS c " & _
               "LEFT OUTER JOIN (SELECT Usuario, Casa, COUNT(*) AS Cuantos " & _
                                  "FROM dbo.discos_Paquetes " & _
                              "GROUP BY Usuario, Casa) AS t " & _
                           "ON c.Usuario = t.Usuario " & _
                          "AND c.Codigo = t.Casa " & _
                        "WHERE (c.Codigo <> '00000000') " & _
                          "AND (c.Usuario = '" & Usuario & "') " 

            if verTipo <> "*" then sqlString = sqlString & vv 
            sqlString = sqlString & "ORDER BY " & oo

            set t = con.execute(sqlString)
        %>          

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 40%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                Casas Editoras
            </div>
            
            <div style="flex: 0 0 60%; text-align: right;">
                <select class="no-field" style="width: 200px;" name="cboVer" id="cboVer" onChange="requery();">
                    <option value="1" <% if verTipo = "1" then response.write " selected" %>>Casas Activas</option>
                    <option value="0" <% if verTipo = "0" then response.write " selected" %>>Casas Desaparecidas</option>
                    <option value="3" <% if verTipo = "3" then response.write " selected" %>>Todas las Casas</option>
                </select>   

                <select class="no-field" style="width: 125px;" name="cboOrdenamiento" id="cboOrdenamiento" onChange="requery();">
                    <option value="1" <% if ordenamiento = "1" then response.write " selected" %>>▲ Nombre</option>
                    <option value="2" <% if ordenamiento = "2" then response.write " selected" %>>▲ Medios</option>
                    <option value="3" <% if ordenamiento = "3" then response.write " selected" %>>▼ Nombre</option>
                    <option value="4" <% if ordenamiento = "4" then response.write " selected" %>>▼ Medios</option>                 
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
                                <th class="sticky" style="width: 65%; text-align: center;">Nombre</th>
                                <th class="sticky" style="width: 10%; text-align: center;">Medios</th>
                                <th class="sticky" style="width: 25%; text-align: center;">&nbsp</th>
                            </tr>
                        </thead>

                        <tbody>
                            <%
                                if not (t.bof or t.eof) then
                                    sw = -1

                                    Do
                                        sw = -1 * sw 
                                        cuantos = cuantos + 1
                                        vinculo = "editar.asp?c=" & t("Codigo") & "&o=" & ordenamiento & "&v=" & verTipo
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

                                                    <td style="text-align: center;">
                                                        <a class="linea" href="<%= vinculo %>">
                                                            <%= t("Medios") %>
                                                        </a>
                                                    </td>

                                                    <td style="text-align:center; padding: 2px;">
                                                        <a href="../../lista_paquetes.asp?f=<%= "CasaDisquera = '" & t("Codigo") & "' " %>&t=<%= "Casa = '" & t("Nombre") & "'" %>" >
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
                                <td class="sticky" colspan="3" style="text-align: center;">
                                    <%
                                        if cuantos = 0 then
                                            response.write "No se ha encontrado ninguna casa editora"
                                        else
                                            response.write "Se han encontrado " & cuantos & " casas editoras"
                                        end if
                                    %>                                  
                                </td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>

            <%
                sqlString = "exec discos_gfx_listaCasas '" & Usuario & "'"
                apexColumns "", "chart", sqlString, "Nombre", "Cuantos", "#3c7fb6ff", 150                 
            %>
        </div>
  
        <br /><br />  

        <script>
            function requery() {
                var ver = document.getElementById("cboVer").value;
                var ordenamiento = document.getElementById("cboOrdenamiento").value;

                var vinculo = "lista.asp?v=" + ver + "&o=" + ordenamiento;
                window.location.href = vinculo;
            }

            function nuevo() {
                var ver = document.getElementById("cboVer").value;
                var ordenamiento = document.getElementById("cboOrdenamiento").value;
                var vinculo = "nuevo.asp?v=" + ver + "&o=" + ordenamiento;

                window.location.href = vinculo;
            }            

            function borrar(codigo) {
                var confirmacion = confirm("Está seguro de borrar el registro seleccionado?");

                var grupo = document.getElementById("cboVer").value;
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