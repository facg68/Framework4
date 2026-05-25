<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Editar Generos</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  

        <%
            thisSystem = "discos"
            thisProcess = "discos.0230"
            SysLockOut

            dim con, t, tt, sqlString, data, labels
            dim cbox, cuantos, ordenamiento, oo, vv, cat, verCat
            dim Codigo, Nombre, Descripcion, Cuenta, vinculo, verTipo   

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")     
        %>      

        <style>
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
            VerCat = Request.QueryString("v")
        
            cuantos = 0
            if ordenamiento = "" then ordenamiento = "1"
            if VerCat = "" then verCat = "1"

            SELECT CASE ordenamiento
                Case 1: oo = "Nombre ASC;"
                Case 2: oo = "Medios ASC;"
                Case 3: oo = "Nombre DESC;"
                Case 4: oo = "Medios DESC;"
            END SELECT  

            '
            ' Comando SQL
            '

            sqlString = "SELECT Usuario, Codigo, Nombre, isnull(dbo.discos_Total_Genero(Usuario, Codigo,"

            Select Case VerCat
                case "0": sqlString = sqlString & "'*'" 
                case "1": sqlString = sqlString & "'DM'"
                case "2": sqlString = sqlString & "'PE'"
                case "3": sqlString = sqlString & "'JU'"
                case "4": sqlString = sqlString & "'SO'"
                case "5": sqlString = sqlString & "'LI'"
                case "6": sqlString = sqlString & "'HW'"
            end select

            sqlString = sqlString & "),0) AS Medios " & _
                            "FROM dbo.discos_Tipos AS t " & _
                            "WHERE (Codigo <> '00000000') " & _
                            "AND (Usuario = '" & Usuario & "') "

            SELECT CASE VerCat
                Case "1": sqlString = sqlString & "AND (Musica = 1) "
                Case "2": sqlString = sqlString & "AND (Video = 1) "
                Case "3": sqlString = sqlString & "AND (Juegos = 1) "
                Case "4": sqlString = sqlString & "AND (Software = 1) "
                Case "5": sqlString = sqlString & "AND (Libros = 1) "
                Case "6": sqlString = sqlString & "AND (Hardware = 1) "
            END SELECT   

            sqlString = sqlString & "ORDER BY "

            SELECT CASE ordenamiento
                Case "1": sqlString = sqlString & "Nombre;"
                Case "2": sqlString = sqlString & "Medios;"
                Case "3": sqlString = sqlString & "Nombre Desc;"
                Case "4": sqlString = sqlString & "Medios Desc;"
            END SELECT             

            set t = con.execute(sqlString)
        %>          

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 40%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                Categorías
            </div>
            
            <div style="flex: 0 0 60%; text-align: right;">
                <select class="no-field" style="width: 125px;" name="cboCateg" id="cboCateg" onChange="requery();">
                    <option value="0" <% if VerCat = "0" then response.write " selected" %>> - - Todo - - </option>                                   
                    <option value="1" <% if VerCat = "1" then response.write " selected" %>>Musica</option>
                    <option value="2" <% if VerCat = "2" then response.write " selected" %>>Video</option>
                    <option value="3" <% if VerCat = "3" then response.write " selected" %>>Juegos</option>
                    <option value="4" <% if VerCat = "4" then response.write " selected" %>>Software</option>                 
                    <option value="5" <% if VerCat = "5" then response.write " selected" %>>Libros</option>                 
                    <option value="6" <% if VerCat = "6" then response.write " selected" %>>Hardware</option>                 
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
                                        vinculo = "editar.asp?c=" & t("Codigo") & "&o=" & ordenamiento & "&v=" & verTipo
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
                                                        <a href="../../lista_paquetes.asp?f=<%= "CHARINDEX('" & t("Codigo") & "', ListaGeneros) > 0 " %>&t=<%=  "Género = '" & t("Nombre") & "'" %>" >
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
                                            response.write "No se ha encontrado ninguna categoría"
                                        else
                                            response.write "Se han encontrado " & cuantos & " categorías"
                                        end if
                                    %>                                  
                                </td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>

            <%
                sqlString = "exec discos_gfx_listaGeneros '" & Usuario & "', " & VerCat
                apexColumns "", "chart", sqlString, "Nombre", "Cuantos", "#3c7fb6ff", 150                 
            %>
        </div>
  
        <br /><br />  

        <script>
            function requery() {
                var ver = document.getElementById("cboCateg").value;
                var ordenamiento = document.getElementById("cboOrdenamiento").value;

                var vinculo = "lista.asp?v=" + ver + "&o=" + ordenamiento;
                window.location.href = vinculo;
            }

            function nuevo() {
                var ver = document.getElementById("cboCateg").value;
                var ordenamiento = document.getElementById("cboOrdenamiento").value;

                var vinculo = "nuevo.asp?v=" + ver + "&o=" + ordenamiento;
                window.location.href = vinculo;
            }            

            function borrar(codigo) {
                var confirmacion = confirm("Está seguro de borrar el registro seleccionado?");

                var ver = document.getElementById("cboCateg").value;
                var ordenamiento = document.getElementById("cboOrdenamiento").value;
                var vinculo = "borrar.asp?c=" + codigo + "&v=" + ver + "&o" + ordenamiento;

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