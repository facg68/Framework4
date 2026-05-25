<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Editar Datos de Video</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "discos"
            thisProcess = "discos.0250"
            SysLockOut
        %>         

        <style>
            body { overflow-y: hidden; }

            a.linea, a.linea:link, a.linea:visited,
            a.linea:focus, a.linea:hover, 
            a.linea:active { color: black; }
        </style>   

        <%
            dim con, t, tt, sqlString, data, labels, cat, verCat, verEstatus
            dim cbox, cuantos, ordenamiento, oo, vv, eest
            dim Codigo, Nombre, Descripcion, Cuenta, vinculo, verTipo   

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")     
        %>           
    </head>

    <body plantilla="dividida" tabla="55" grafica="45" reserva="200">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            Usuario = Request.Cookies("usuario")
            cuantos = 0

            sqlString = "SELECT fp.Usuario, fp.Codigo, fp.Nombre, ISNULL(m.Cuantos, 0) AS Medios " & _
                        "FROM dbo.discos_FormatosPantalla AS fp " & _
                        "LEFT OUTER JOIN (SELECT Usuario, FormatoPantalla, COUNT(*) AS Cuantos " & _
                                        "FROM dbo.discos_Objetos " & _
                                        "GROUP BY Usuario, FormatoPantalla) AS m " & _
                        "ON fp.Usuario = m.Usuario " & _
                        "AND fp.Codigo = m.FormatoPantalla " & _
                        "WHERE (fp.Usuario = '" & Usuario & "') " & _
                        "AND (fp.Codigo <> '00000000') " & _
                        "ORDER BY fp.Nombre;"

            set t = con.execute(sqlString)
        %>         

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 60%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                Formatos de Pantalla
            </div>
            
            <div style="flex: 0 0 30%; text-align: right;">
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
                                        classLinea = ""
                                        vinculo = "editar.asp?c=" & t("Codigo")

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
                                                            <%= FormatNumber(t("Medios"), 0) %>
                                                        </a>
                                                    </td>

                                                    <td style="text-align:center; padding: 2px;">
                                                        <% if classLinea = "disabled" then %>
                                                            <button class="form-btn rojo disabled">
                                                                <i class="fa fa-trash fa-xl" title="Borrar"></i>
                                                            </button>
                                                        <% else %>
                                                            <a onclick="borrar('<%= t("Codigo") %>','<%= t("Nombre") %>')">
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
                                            response.write "No se ha encontrado ningun formato de pantalla"
                                        else
                                            response.write "Se han encontrado " & cuantos & " formatos de pantalla"
                                        end if
                                    %>                                  
                                </td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>

            <%
                sqlString = "exec discos_gfx_listaPantallas '" & Usuario & "'"
                apexColumns "", "chart", sqlString, "Nombre", "Medios", "#3c7fb6ff", 150                 
            %>
        </div>
  
        <br /><br />  

        <script>
            function requery() {
                //var vinculo = "lista.asp?e=" + est + "&v=" + ver + "&o=" + ordenamiento;
                //window.location.href = vinculo;
            }

            function nuevo() {
                var vinculo = "nuevo.asp";
                window.location.href = vinculo;
            }            

            function borrar(codigo) {
                var confirmacion = confirm("Está seguro de borrar el registro seleccionado?");                
                var vinculo = "borrar.asp?c=" + codigo;

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