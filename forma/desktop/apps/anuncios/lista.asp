<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Publicaciones</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->    
                
        <%
            thisSystem = "anuncios"
            thisProcess = "anu.0100"
            SysLockOut

            dim cc, t, tt, sqlString, data, labels
            dim cActivas, cInactivas, estatusAnuncio, ordenadoPor

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn") 

            '
            ' Funciones
            '

            function NombrePantalla(codigoPantalla)
                dim pp 

                set pp = cc.execute("SELECT Nombre FROM seg_Anuncios_Pantallas WHERE Pantalla = '" & codigoPantalla & "';")

                if not (pp.bof or pp.eof) then
                    NombrePantalla = "(" & pp("Nombre") & ")"
                else
                    NombrePantalla = ""
                end if

                pp.close: set pp = nothing
            end function

            function fechaFormulario(FechaSQL)
                dim d, m, a, h, mm

                d = RIGHT("00" & DAY(FechaSQL), 2)
                m = RIGHT("00" & MONTH(FechaSQL), 2)
                a = YEAR(FechaSQL)

                h = RIGHT("00" & HOUR(FechaSQL), 2)
                mm = RIGHT("00" & MINUTE(FechaSQL), 2)

                fechaFormulario = d & "/" & m & "/" & a & " " & h & ":" & m
            end function           
        %>    
    </head>

    <body plantilla="lista" reserva="190">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    

        <br />

        <%
            verPantalla = request.querystring("tv")
            estatusAnuncio = request.querystring("e")
            ordenadoPor = request.querystring("op")            

            if verPantalla = "" then verPantalla = "*"
            if estatusAnuncio = "" then estatusAnuncio = "a"
            if ordenadoPor = "" then ordenadoPor = "0"

            cActivas = 0 
            cInactivas = 0

            if verPantalla <> "*" then
                sqlString = "SELECT DISTINCT Secuencia, Propietario, Titulo, Inicio, Fin, Tipo, VerComo, NombreImagen, Estado, " & _
                                           " EscalarImagen, Cuerpo, Segundos, Orden, Codigo, " & _
                                           " dbo.seg_Anuncios_Up_Sec(Propietario, Orden) AS UpSec, dbo.seg_Anuncios_Up(Propietario, Orden) AS UpOrden, " & _
                                           " dbo.seg_Anuncios_Down_Sec(Propietario, Orden) AS DownSec, dbo.seg_Anuncios_Down(Propietario, Orden) AS DownOrden " & _
                              "FROM (SELECT a.Secuencia, a.Propietario, a.Titulo, a.Inicio, a.Fin, a.Tipo, a.VerComo, a.NombreImagen, a.EscalarImagen, " & _
                                          " a.Cuerpo, a.Segundos, a.Orden, a.Codigo, aa.Pantalla, CASE WHEN Fin >= GETDATE() THEN 1 ELSE 0 END AS Estado " & _
                                      "FROM dbo.seg_Anuncios AS a INNER JOIN dbo.seg_Anuncios_Asignaciones AS aa ON a.Secuencia = aa.Anuncio " & _
                                     "WHERE (aa.Pantalla IN (SELECT Pantalla FROM dbo.seg_Anuncios_Pantallas_Usuarios " & _
                                                             "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') AND (Pantalla = '" & verPantalla & "')))) AS anUsuario " & _
                             "WHERE (Propietario = '" & Request.Cookies("Usuario") & "') "

            else
                sqlString = "SELECT DISTINCT Secuencia, Propietario, Titulo, Inicio, Fin, Tipo, VerComo, NombreImagen, Estado, " & _
                                           " EscalarImagen, Cuerpo, Segundos, Orden, Codigo, " & _
                                           " dbo.seg_Anuncios_Up_Sec(Propietario, Orden) AS UpSec, dbo.seg_Anuncios_Up(Propietario, Orden) AS UpOrden, " & _
                                           " dbo.seg_Anuncios_Down_Sec(Propietario, Orden) AS DownSec, dbo.seg_Anuncios_Down(Propietario, Orden) AS DownOrden " & _
                              "FROM (SELECT a.Secuencia, a.Propietario, a.Titulo, a.Inicio, a.Fin, a.Tipo, a.VerComo, a.NombreImagen, a.EscalarImagen, " & _
                                          " a.Cuerpo, a.Segundos, a.Orden, a.Codigo, aa.Pantalla, CASE WHEN Fin >= GETDATE() THEN 1 ELSE 0 END AS Estado " & _
                                      "FROM dbo.seg_Anuncios AS a INNER JOIN dbo.seg_Anuncios_Asignaciones AS aa ON a.Secuencia = aa.Anuncio " & _
                                     "WHERE (aa.Pantalla IN (SELECT Pantalla FROM dbo.seg_Anuncios_Pantallas_Usuarios " & _
                                                             "WHERE (Usuario = '" & Request.Cookies("Usuario") & "')))) AS anUsuario " & _
                             "WHERE (Propietario = '" & Request.Cookies("Usuario") & "') "                                    
            end if         

            select case estatusAnuncio
                case "a": sqlString = sqlString & "AND (Estado = 1) "
                case "h": sqlString = sqlString & "AND (Estado = 1) AND (GETDATE() BETWEEN Inicio AND Fin) "
                case "d": sqlString = sqlString & "AND (Estado = 0) "
            end select            

            select case ordenadoPor
                case 0: sqlString = sqlString & " ORDER BY Orden;"
                case 1: sqlString = sqlString & " ORDER BY Titulo;"
                case 2: sqlString = sqlString & " ORDER BY Titulo Desc;"
                case 3: sqlString = sqlString & " ORDER BY Inicio;"
                case 4: sqlString = sqlString & " ORDER BY Inicio Desc;"
                case 5: sqlString = sqlString & " ORDER BY Fin;"
                case 6: sqlString = sqlString & " ORDER BY Fin Desc;"
            end select

            set t = cc.execute(sqlString)        
        %>

        <div style="width: 98%; margin: auto;">
            <table style="width: 100%; margin: auto;">
                <tr class="noborder" style="padding: 10px;">
                    <td colspan="2" style="text-align:left; width: 20%;">
                        <span style="font-size: 24px">
                            &nbsp;
                            <%
                                select case estatusAnuncio
                                    case "a": response.write "Publicaciones Activas"
                                    case "h": response.write "Publicaciones de Hoy"
                                    case "d": response.write "Publicaciones Desctivadas"
                                    case else: response.write "Publicaciones"
                                end select   

                                if NombrePantalla(verPantalla) <> "" then
                                    response.write "&nbsp;" &  NombrePantalla(verPantalla)
                                end if
                            %>
                        </span>

                        <br />

                        <span style="font-size: 20px">
                            &nbsp;                        
                            <%
                                select case ordenadoPor
                                    case 0: response.write "&nbsp;En Orden de Aparición"
                                    case 1: response.write "&nbsp;Ordenado por Título"
                                    case 2: response.write "&nbsp;Ordenado por Título (descendentemente)"
                                    case 3: response.write "&nbsp;Ordenado por Fecha de Inicio"
                                    case 4: response.write "&nbsp;Ordenado por Fecha de Inicio (descendentemente)"
                                    case 5: response.write "&nbsp;Ordenado por Fecha de Finalización"
                                    case 6: response.write "&nbsp;Ordenado por Fecha de Finalización (descendentemente)"
                                end select                            
                            %>
                        </span>
                    </td>

                    <td colspan="3" style="text-align:right; width: 20%;">
                        <%
                            sqlString = "SELECT up.Pantalla, p.Nombre " & _
                                          "FROM dbo.seg_Anuncios_Pantallas_Usuarios AS up " & _
                                    "INNER JOIN dbo.seg_Anuncios_Pantallas AS p " & _
                                            "ON up.Pantalla = p.Pantalla " & _
                                         "WHERE (up.Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                      "ORDER BY p.Nombre;"

                            set l = cc.execute(sqlString)

                            if not (l.bof or l.eof) then
                                swSubmit = true

                                %>
                                    <select class="field" name="codpantalla" id="codpantalla" required 
                                            onChange="filtrar()">
                                
                                    <option value="*" <% if verPantalla = "*" then response.write " selected" %>>- - Todas las Pantallas - -</option>
                                <%

                                do
                                    response.write "<option value='" & l("Pantalla") & "' "
                                        if l("Pantalla") = verPantalla then
                                            response.write " selected"
                                        end if
                                    response.write ">" & l("Nombre") & "</option>"
                                    l.movenext
                                loop until l.eof

                                response.write "</select>"
                            end if

                            l.close: set l = nothing
                        %>    

                        <select class="field" name="ordenadoPor" id="ordenadoPor" required 
                                onChange="filtrar()">
                            <option value="0" <% if ordenadoPor = "0" then response.write " selected" %>>Presentacion</option>
                            <option value="1" <% if ordenadoPor = "1" then response.write " selected" %>>&#9650; Titulo</option>
                            <option value="2" <% if ordenadoPor = "2" then response.write " selected" %>>&#9660; Titulo</option>
                            <option value="3" <% if ordenadoPor = "3" then response.write " selected" %>>&#9650; Desde</option>
                            <option value="4" <% if ordenadoPor = "4" then response.write " selected" %>>&#9660; Desde</option>
                            <option value="5" <% if ordenadoPor = "5" then response.write " selected" %>>&#9650; Hasta</option>
                            <option value="6" <% if ordenadoPor = "6" then response.write " selected" %>>&#9660; Hasta</option>
                        </select>                        

                        <select class="field" name="verlista" id="verlista" required 
                                onChange="filtrar()">
                            <option value="*" <% if estatusAnuncio = "*" then response.write " selected" %>>Ver Todo</option>
                            <option value="a" <% if estatusAnuncio = "a" then response.write " selected" %>>Activas</option>
                            <option value="h" <% if estatusAnuncio = "h" then response.write " selected" %>>Hoy (Activas)</option>
                            <option value="d" <% if estatusAnuncio = "d" then response.write " selected" %>>Desactivadas</option>
                        </select>                                           

                        <button type="button" class="form-btn verde" onclick="editar('*')">
                            <i class=" fa fa-edit fa-xl" title="Nuevo"></i>
                        </button>                        
                    </td>                    
                </tr>
            </table>

            <table style="width:98%;">
                <tr style="font-size: 14px; background-color: rgb(61, 61, 61); color:rgb(255,255,255);">
                    <td style="padding: 10px; text-align:center; width: 10%;">Estado</td>
                    <td style="padding: 10px; text-align:center; width: 40%;">Titulo</td>
                    <td style="padding: 10px; text-align:center; width: 15%;">Desde</td>
                    <td style="padding: 10px; text-align:center; width: 15%;">Hasta</td>
                    <td style="padding: 10px; text-align:center; width: 25%;">&nbsp;</td>
                </tr>

                <tr>
                    <td colspan="5">
                        <div id="overFlow" style="width:100%; height: 650px; overflow: auto; background-color: rgb(207, 207, 207);">
                            <table style="width: 100%;">
                                <%  
                                    if not (t.bof or t.eof) then  
                                        Do     
                                            bgcolor = "199, 230, 188"
                                            if t("Estado") = "0" then bgcolor = "245, 211, 208"
                                %>
                                        <tr style="font-size: 14px; background-color: rgb(255,255,255); color:rgb(0,0,0); border-bottom: 1px solid rgb(194, 194, 194);" >
                                            <td style="padding: 10px; text-align:center; width: 10%; background-color: rgb(<%= bgcolor %>);" onclick="editar(<%= t("secuencia") %>)">
                                                <%
                                                    select case t("Estado")
                                                        case "1"
                                                            response.write "Activa"
                                                            cActivas = cActivas + 1
                                                        case "0"
                                                            response.write "Inactiva"
                                                            cInactivas = cInactivas + 1                                                        
                                                    end select                                                
                                                %>
                                            </td>


                                            <td style="padding: 5px; text-align:center; width: 35%;" onclick="editar(<%= t("secuencia") %>)">
                                                <%
                                                    response.write t("Titulo")
                                                %>
                                            </td>

                                            <%
                                                if (ordenadoPor = 0) and (estatusAnuncio = "a") then
                                                    if isnull(t("upSec")) then
                                                        %>
                                                            <td style="padding: 10px; text-align:center; width: 2%;">
                                                                &nbsp;
                                                            </td>
                                                        <%
                                                    else
                                                        %>
                                                            <td style="padding: 10px; text-align:center; width: 2%;" onclick="intercambiar(<%= t("secuencia") %>, <%= t("orden") %>, <%= t("UpSec") %>, <%= t("UpOrden") %>)">
                                                                <img src="imagenes/up.png" style = "border: none;">
                                                            </td>
                                                        <%
                                                    end if
                                                else
                                                    %>
                                                        <td style="padding: 10px; text-align:center; width: 2%;">
                                                            &nbsp;
                                                        </td>
                                                    <%                                                
                                                end if
                                            %>

                                            <td style="padding: 10px; text-align:center; width: 1%;">&nbsp;</td>

                                            <%
                                                if (ordenadoPor = 0) and (estatusAnuncio = "a") then
                                                    if isnull(t("DownSec")) then
                                                        %>
                                                            <td style="padding: 10px; text-align:center; width: 2%;">
                                                                &nbsp;
                                                            </td>
                                                        <%
                                                    else
                                                        %>
                                                            <td style="padding: 10px; text-align:center; width: 3%;" onclick="intercambiar(<%= t("secuencia") %>, <%= t("orden") %>, <%= t("DownSec") %>, <%= t("DownOrden") %>)">
                                                                <img src="imagenes/down.png" style = "border: none;">
                                                            </td>
                                                        <%
                                                    end if
                                                else
                                                    %>
                                                        <td style="padding: 10px; text-align:center; width: 2%;">
                                                            &nbsp;
                                                        </td>
                                                    <%                                                
                                                end if                                                    
                                            %>

                                            <td style="padding: 5px; text-align:center; width: 15%;" onclick="editar(<%= t("secuencia") %>)"><%= fechaFormulario(t("Inicio")) %></td>
                                            <td style="padding: 5px; text-align:center; width: 15%;" onclick="editar(<%= t("secuencia") %>)"><%= fechaFormulario(t("Fin")) %></td>
                                            <td style="padding: 5px; text-align:right; width:20%;">
                                                <button type="button" class="form-btn azul" onclick="republicar(<%= t("secuencia") %>)">
                                                    <i class=" fa fa-copy fa-xl" title="Volver a Publicar"></i>
                                                </button>

                                                <button type="button" class="form-btn verde" onclick="ver(<%= t("secuencia") %>)">
                                                    <i class=" fa fa-eye fa-xl" title="Ver Publicacion"></i>
                                                </button>

                                                <% if t("tipo") = 1 or t("tipo") = 3 then %>
                                                    <button type="button" class="form-btn violeta" onclick="imagen(<%= t("secuencia") %>)">
                                                        <i class=" fa fa-cloud fa-xl" title="Subir Objeto"></i>
                                                    </button>
                                                <% else %>
                                                    <button type="button" class="form-btn violeta disabled" disabled>
                                                        <i class=" fa fa-cloud fa-xl" title="Subir Objeto"></i>
                                                    </button>                                                
                                                <% end if %>

                                                <button type="button" class="form-btn rojo" onclick="borrar(<%= t("secuencia") %>)">
                                                    <i class=" fa fa-trash fa-xl" title="Borrar Publicacion"></i>
                                                </button>
                                            </td>
                                        </tr>
                                <% 
                                            t.MoveNext
                                        Loop Until (t.eof)
                                    end if 
                                %>
                            </table>
                        </div>                
                    </td>
                </tr>

                <tr style="font-size: 14px; background-color: rgb(61, 61, 61); color:rgb(255,255,255);">
                    <td colspan="12" style="padding: 10px; text-align:center; width: 100%;">
                        &nbsp;&nbsp;Activas:&nbsp;<%= cActivas %>&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp;
                        &nbsp;&nbsp;Inactivas:&nbsp;<%= cInactivas %>&nbsp;&nbsp;
                    </td>
                </tr>                               
            </table>
        </div>

        <%
            t.close: set t = nothing
        %>

        <script>
            function filtrar() {
                var pantalla = document.getElementById("codpantalla").value;
                var estatus = document.getElementById("verlista").value;
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="lista.asp?tv=" + pantalla + "&e=" + estatus + "&op=" + ordenamiento;
                window.location.href = vinculo;                      
            }
                  
            function editar(publicacion) {
                var pantalla;

                var el = document.getElementById("codpantalla");

                if (!el || el.value.trim() === "") {
                    pantalla = "";
                } else {
                    pantalla = el.value;
                }
                
                var estatus = document.getElementById("verlista").value;
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="editar.asp?w=0&a=" + publicacion + "&tv=" + pantalla + "&e=" + estatus + "&op=" + ordenamiento;
                window.location.href = vinculo;
            }    

            function republicar(publicacion) {
                var pantalla = document.getElementById("codpantalla").value;
                var estatus = document.getElementById("verlista").value;
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="editar.asp?w=0&a=" + publicacion + "&tv=" + pantalla + "&e=" + estatus + "&op=" + ordenamiento + "&r=1";
                window.location.href = vinculo;
            }                
            
            function intercambiar(secuencia1, orden1, secuencia2, orden2) {
                var pantalla = document.getElementById("codpantalla").value;
                var estatus = document.getElementById("verlista").value;
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="mover.asp?w=0&s1=" + secuencia1 + "&o1=" + orden1 + "&s2=" + secuencia2 + "&o2=" + orden2 + "&tv=" + pantalla + "&e=" + estatus + "&op=" + ordenamiento;
                window.location.href = vinculo; 
            }      

            function borrar(publicacion) {
                var confirmacion = confirm("Desea borrar la publicación seleccionada?");
                var pantalla = document.getElementById("codpantalla").value;
                var estatus = document.getElementById("verlista").value;
                var ordenamiento = document.getElementById("ordenadoPor").value;
                var vinculo ="borrar.asp?w=0&a=" + publicacion + "&tv=" + pantalla + "&e=" + estatus + "&op=" + ordenamiento;                

                if (confirmacion) {     
                    window.location.href = vinculo;
                };
            }    

            function imagen(publicacion) {
                var pantalla = document.getElementById("codpantalla").value;
                var estatus = document.getElementById("verlista").value;
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="subir_objeto.asp?w=0&a=" + publicacion + "&tv=" + pantalla + "&e=" + estatus + "&op=" + ordenamiento;                                         
                window.location.href = vinculo;          
            }            
                    
            function ver(publicacion) {
                var pantalla = document.getElementById("codpantalla").value;
                var estatus = document.getElementById("verlista").value;
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="ver.asp?w=0&a=" + publicacion + "&tv=" + pantalla + "&e=" + estatus + "&op=" + ordenamiento;                                         
                window.location.href = vinculo;          
            }       
        </script> 

        <% cc.close: set cc = nothing %>
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>