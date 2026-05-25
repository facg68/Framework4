<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Editar Colecciones</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  

        <%
            thisSystem = "discos"
            thisProcess = "discos.0260"
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

    <body plantilla="dividida" tabla="55" grafica="45" reserva="200">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            Usuario = Request.Cookies("usuario")
            ordenamiento = Request.QueryString("o")
        
            cuantos = 0
            if ordenamiento = "" then ordenamiento = "1"

            SELECT CASE ordenamiento
                Case 1: oo = "Nombre ASC;"
                Case 2: oo = "Medios ASC;"
                Case 3: oo = "Nombre DESC;"
                Case 4: oo = "Medios DESC;"
            END SELECT  

            sqlString = "SELECT c.Usuario, c.Codigo, c.Nombre, c.Descripcion, c.PorDefecto, c.DeSistema, ISNULL(m.Cuantos, 0) AS Medios " & _
                        "FROM dbo.discos_Carpetas AS c " & _
                        "LEFT OUTER JOIN (SELECT Usuario, Carpeta, COUNT(*) AS Cuantos " & _
                                        "FROM dbo.discos_Paquetes " & _
                                        "GROUP BY Usuario, Carpeta) AS m " & _
                        "ON (c.Usuario = m.Usuario) " & _
                        "AND (c.Codigo = m.Carpeta) " & _
                        "WHERE (c.Usuario = '" & Request.Cookies("Usuario") & "') " 

            sqlString = sqlString & "ORDER BY " & oo

            set t = con.execute(sqlString)
        %>         

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 60%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                Colecciones
            </div>
            
            <div style="flex: 0 0 30%; text-align: right;">
                &nbsp;Orden&nbsp;

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
                                <th class="sticky" style="width: 75%; text-align: center;">Nombre</th>
                                <th class="sticky" style="width: 10%; text-align: center;">Medios</th>
                                <th class="sticky" style="width: 15%; text-align: center;">&nbsp</th>
                            </tr>
                        </thead>

                        <tbody>
                            <%
                                if not (t.bof or t.eof) then
                                    sw = -1

                                    Do
                                        sw = -1 * sw 
                                        cuantos = cuantos + 1
                                        vinculo = "editar.asp?c=" & t("Codigo") & "&o=" & ordenamiento 
                                        classLinea = ""

                                        if cint(t("Medios")) > 0 then
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
                                            response.write "No se ha encontrado ninguna colección"
                                        else
                                            response.write "Se han encontrado " & cuantos & " colecciones"
                                        end if
                                    %>                                  
                                </td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>

            <%
                sqlString = "exec discos_gfx_listaFolders '" & Usuario & "'"
                apexColumns "", "chart", sqlString, "Nombre", "Cuantos", "#3c7fb6ff", 150                 
            %>
        </div>
  
        <br /><br />  

        <script>
            function requery() {
                var ordenamiento = document.getElementById("cboOrdenamiento").value;
                var vinculo = "lista.asp?o=" + ordenamiento;
                window.location.href = vinculo;
            }

            function nuevo() {
                var ordenamiento = document.getElementById("cboOrdenamiento").value;
                var vinculo = "nuevo.asp?o=" + ordenamiento;
                window.location.href = vinculo;
            }            

            function borrar(codigo) {
                var confirmacion = confirm("Está seguro de borrar el registro seleccionado?");

                var ordenamiento = document.getElementById("cboOrdenamiento").value;
                var vinculo = "borrar.asp?c=" + codigo + "&o=" + ordenamiento;

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