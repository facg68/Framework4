<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Editar Plataformas</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "discos"
            thisProcess = "discos.0240"
            SysLockOut
        %>         

        <style>
            a.linea, a.linea:link, a.linea:visited,
            a.linea:focus, a.linea:hover, 
            a.linea:active { color: black; }
        </style>   

        <%
            dim con, t, tt, sqlString, data, labels, cat, VerTipo, verEstatus
            dim cbox, cuantos, ordenamiento, oo, vv, eest
            dim Codigo, Nombre, Descripcion, Cuenta, vinculo   

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")     
        %>           
    </head>

    <body plantilla="dividida" tabla="55" grafica="45" reserva="200">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            Usuario = Request.Cookies("usuario")
            ordenamiento = Request.QueryString("o")
            VerTipo = Request.QueryString("t")
            verEstatus = Request.QueryString("e")
        
            cuantos = 0
            if ordenamiento = "" then ordenamiento = "1"
            if VerTipo = "" then VerTipo = "*"
            if verEstatus = "" then verEstatus = "0"

            SELECT CASE ordenamiento
                Case 1: oo = "Nombre ASC;"
                Case 2: oo = "Medios ASC;"
                Case 3: oo = "Nombre DESC;"
                Case 4: oo = "Medios DESC;"
            END SELECT  

            SELECT CASE VerTipo
                Case "1": cat = "AND (Juegos = 1) "
                Case "2": cat = "AND (Software = 1) "
            END SELECT   

            SELECT CASE verEstatus
                Case "0": eest = "AND (Obsoleta = 0) "
                Case "1": eest = "AND (Obsoleta = 1) "
            END SELECT          

            sqlString = "SELECT p.Usuario, p.Codigo, p.Nombre, p.Software, p.Juegos, p.Obsoleta, ISNULL(m.Cuantos, 0) AS Medios " & _
                        "FROM dbo.discos_Plataformas AS p " & _
                        "LEFT OUTER JOIN (SELECT Usuario, PlatOS, COUNT(*) AS Cuantos " & _
                                        "FROM dbo.discos_Objetos " & _
                                    "GROUP BY Usuario, PlatOS) AS m " & _
                        "ON p.Usuario = m.Usuario " & _
                        "AND p.Codigo = m.PlatOS " & _
                        "WHERE (p.Codigo <> '00000000') " & _
                        "AND (p.Usuario = '" & Request.Cookies("Usuario") & "') " 

            if VerTipo <> "*" then sqlString = sqlString & cat
            if verEstatus <> "*" then sqlString = sqlString & eest

            sqlString = sqlString & "ORDER BY " & oo

            set t = con.execute(sqlString)
        %>         

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 30%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                Plataformas
            </div>
            
            <div style="flex: 0 0 70%; text-align: right;">
                <select class="no-field" style="width: 125px;" name="cboEstatus" id="cboEstatus" onChange="requery();">
                    <option value="*" <% if verEstatus = "*" then response.write " selected" %>> - - Todo - - </option>                                   
                    <option value="0" <% if verEstatus = "0" then response.write " selected" %>>Activas</option>
                    <option value="1" <% if verEstatus = "1" then response.write " selected" %>>Obsoletas</option>                 
                </select>  

                <select class="no-field" style="width: 125px;" name="cboCateg" id="cboCateg" onChange="requery();">
                    <option value="*" <% if VerTipo = "*" then response.write " selected" %>> - - Todo - - </option>                                   
                    <option value="1" <% if VerTipo = "1" then response.write " selected" %>>Juegos</option>
                    <option value="2" <% if VerTipo = "2" then response.write " selected" %>>Software</option>                 
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
                                        vinculo = "editar.asp?c=" & t("Codigo") & "&t=" & VerTipo & "&e=" & verEstatus & "&o=" & Ordenamiento
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
                                                        <a href="../../lista_paquetes.asp?f=<%= "CHARINDEX('" & t("Codigo") & "', ListaPlataformas) > 0 " %>&t=<%= "Plataforma = '" & t("Nombre") & "'" %>" >
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
                                            response.write "No se ha encontrado ninguna plataforma"
                                        else
                                            response.write "Se han encontrado " & cuantos & " plataformas"
                                        end if
                                    %>                                  
                                </td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>

            <%
                sqlString = "exec discos_gfx_listaPlataformas '" & Usuario & "'"
                apexColumns "", "chart", sqlString, "Nombre", "Cuantos", "#3c7fb6ff", 150                 
            %>
        </div>
  
        <br /><br />  

        <script>
            function requery() {
                var est = document.getElementById("cboEstatus").value;
                var ver = document.getElementById("cboCateg").value;
                var ordenamiento = document.getElementById("cboOrdenamiento").value;

                var vinculo = "lista.asp?e=" + est + "&t=" + ver + "&o=" + ordenamiento;
                window.location.href = vinculo;
            }

            function nuevo() {
                var est = document.getElementById("cboEstatus").value;
                var ver = document.getElementById("cboCateg").value;
                var ordenamiento = document.getElementById("cboOrdenamiento").value;

                var vinculo = "nuevo.asp?e=" + est + "&t=" + ver + "&o=" + ordenamiento;
                window.location.href = vinculo;
            }            

            function borrar(codigo) {
                var confirmacion = confirm("Está seguro de borrar el registro seleccionado?");

                var est = document.getElementById("cboEstatus").value;
                var ver = document.getElementById("cboCateg").value;
                var ordenamiento = document.getElementById("cboOrdenamiento").value;
                
                var vinculo = "borrar.asp?c=" + codigo + "&e=" + est + "&t=" + ver + "&o=" + ordenamiento;

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