<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->   
        <%
            thisSystem = "discos"
            thisProcess = "discos.0110"
            SysLockOut

            function CarpetaPorDefecto()
                dim cc, tt

                sqlString = "SELECT Codigo FROM discos_Carpetas WHERE Usuario = '" & Request.Cookies("Usuario") & "' AND PorDefecto = 1;"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.execute(sqlString)

                    CarpetaPorDefecto = tt("Codigo")

                tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function         
        %>      

        <style>
            a.linea:link,
            a.linea:visited,
            a.linea:focus,
            a.linea:hover,
            a.linea:active {
                color: black !important;
            }
            
            td { 
                padding: 2 !important;
                vertical-align: middle !important;
            }
        </style>                 
    </head>

    <body plantilla="tabla" tabla="100" reserva="135">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->
        <%       
            '
            ' Abrimos la tabla y llenamos los datos
            '

            dim con, t, tt, sqlString, usu, vinculo, sw
            dim cuantos, objetos, total, aActual, lsqlString
            dim tipo, forma, amo, folder, ordenamiento

            usu = Request.Cookies("Usuario")

            folder = Request.Cookies("lista_medios")("folder") 
            tipo = Request.Cookies("lista_medios")("tipo") 
            forma = Request.Cookies("lista_medios")("forma") 
            plataforma = Request.Cookies("lista_medios")("plataforma") 
            amo = Request.Cookies("lista_medios")("amo") 
            ordenamiento = Request.Cookies("lista_medios")("ordenamiento")

            if folder = "" then folder = CarpetaPorDefecto()
            if tipo = "" then tipo = "*"
            if forma = "" then forma = "*"
            if plataforma = "" then plataforma = "*"
            if ordenamiento = "" then ordenamiento = "6"
            if amo = "" then amo = 10
            aActual = Year(Now())

            cuantos = 0
            objetos = 0
            total = 0.00

            '
            ' Creamos la cadena de conexión, dependiendo de los
            ' datos del filtro, o generamos una cadena nueva

            sqlString = "SELECT Usuario, Paquete, Amo, AmoCompra, Titulo, Casa, Medios, Forma, Precio, Coleccion, Medio3D, " & _
                            " ListaFormas, ListaTipos, ListaInDirAu, ListaPlataformas, Icono_Forma " & _
                        "FROM discos_MediosPorForma " & _
                        "WHERE (Usuario = '" & usu & "') " & _
                        "AND (Coleccion = '" & folder & "') " 

            if tipo <> "*"  then sqlString = sqlString & "AND (CHARINDEX('" & tipo & "', ListaTipos) > 0) "
            if forma <> "*" then sqlString = sqlString & "AND (CHARINDEX('" & forma & "', ListaFormas) > 0) "        
            if (tipo = "JU") or (Tipo = "SO")  then 
                if plataforma <> "*" then
                    sqlString = sqlString & "AND (CHARINDEX('" & Plataforma & "', ListaPlataformas) > 0) "
                end if
            end if  

            select case amo
                case 1: sqlString = sqlString & "AND (Amo = " & aActual & ") " 
                case 2: sqlString = sqlString & "AND (Amo = " & (aActual - 1) & ") " 
                case 3: sqlString = sqlString & "AND (Amo BETWEEN " & (aActual - 1) & " AND " & aActual & ") " 
                case 4: sqlString = sqlString & "AND (Amo BETWEEN " & (aActual - 4) & " AND " & aActual & ") " 
                case 5: sqlString = sqlString & "AND (Amo BETWEEN " & (aActual - 9) & " AND " & aActual & ") " 
                case 6: sqlString = sqlString & "AND (Amo BETWEEN " & (aActual - 14) & " AND " & aActual & ") " 
                case 7: sqlString = sqlString & "AND (Amo BETWEEN " & (aActual - 19) & " AND " & aActual & ") " 

                case 8: sqlString = sqlString & "AND (AmoCompra = " & aActual & ") " 
                case 9: sqlString = sqlString & "AND (AmoCompra = " & (aActual - 1) & ") " 
                case 10: sqlString = sqlString & "AND (AmoCompra BETWEEN " & (aActual - 1) & " AND " & aActual & ") " 
                case 11: sqlString = sqlString & "AND (AmoCompra BETWEEN " & (aActual - 4) & " AND " & aActual & ") " 
                case 12: sqlString = sqlString & "AND (AmoCompra BETWEEN " & (aActual - 9) & " AND " & aActual & ") " 
                case 13: sqlString = sqlString & "AND (AmoCompra BETWEEN " & (aActual - 14) & " AND " & aActual & ") " 
                case 14: sqlString = sqlString & "AND (AmoCompra BETWEEN " & (aActual - 19) & " AND " & aActual & ") " 
            end select

            SqlString = sqlString & "ORDER BY "

            SELECT CASE ordenamiento
                CASE 1: sqlString = sqlString & "Amo, Paquete;"
                CASE 2: sqlString = sqlString & "Titulo;" 
                CASE 3: sqlString = sqlString & "Casa;"
                CASE 4: sqlString = sqlString & "Medios;"
                CASE 5: sqlString = sqlString & "Precio;"

                CASE 6: sqlString = sqlString & "Amo DESC, Paquete Desc;"
                CASE 7: sqlString = sqlString & "Titulo DESC;"
                CASE 8: sqlString = sqlString & "Casa DESC;"
                CASE 9: sqlString = sqlString & "Medios DESC;"
                CASE 10: sqlString = sqlString & "Precio DESC;"
            END SELECT

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")
            set t = con.Execute(sqlString)        
        %>        

        <br />

        <form id="formulario" name="formulario" method="post" action="listafiltro.asp">
            <div style="display: flex; justify-content: space-between; width: 93%; margin: auto;">
                <div style="flex: 0 0 90%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    <select class="no-field" name="cboFolder" id="cboFolder" onChange="Requery();">
                        <%
                            sqlString = "Select Codigo, Nombre from discos_Carpetas WHERE Usuario = '" & usu & "' ORDER BY Nombre;"
                            set tt = con.execute(sqlString)

                            if not (tt.bof or tt.eof) then
                                Do
                                response.write "<option value='" & tt("Codigo") & "'"
                                    if Folder = tt("Codigo") then response.write " selected" 
                                response.write ">" & tt("Nombre") & "</option>"

                                tt.MoveNext
                                Loop Until tt.eof
                            end if

                            tt.close: set tt = nothing
                        %>
                    </select>          

                    <select class="no-field" name="cboTipo" id="cboTipo" onChange="Requery();">
                        <option value="*" <% if Forma = "*"   then response.write " selected" %>>- - Tipo - -</option>
                        <%
                            set tt = con.execute("select Codigo, Nombre from discos_Objetos_Clases ORDER BY Nombre ")

                            if not (tt.bof or tt.eof) then
                            Do
                                response.write "<option value='" & tt("Codigo") & "'"
                                if Tipo = tt("Codigo") then response.write " selected" 
                                response.write ">" & tt("Nombre") & "</option>"

                                tt.MoveNext
                            Loop Until tt.eof
                            end if

                            tt.close: set tt = nothing
                        %>
                    </select>   

                    <% if (Tipo = "JU") OR (Tipo = "SO") then %>
                        <select class="no-field" name="cboPlataforma" id="cboPlataforma" onChange="Requery();">
                            <option value="*" <% if Forma = "*"   then response.write " selected" %>>- - Plataforma - -</option>
                                <%
                                    lsqlString = "select Codigo, Nombre " & _
                                                    "from discos_Plataformas " & _ 
                                                    "where (usuario = '" & usu & "') " & _
                                                    "and (Codigo <> '00000000') " 

                                    SELECT CASE Tipo
                                        CASE "JU"
                                            lsqlString = lsqlString & "and (Juegos = 1) "
                                        CASE "SO"
                                            lsqlString = lsqlString & "and (Software = 1) "
                                    END SELECT

                                    lsqlString = lsqlString & "order by Nombre"

                                    set tt = con.execute(lsqlString)

                                    if not (tt.bof or tt.eof) then
                                        Do
                                            response.write "<option value='" & tt("Codigo") & "'"
                                                if Plataforma = tt("Codigo") then response.write " selected" 
                                            response.write ">" & tt("Nombre") & "</option>"

                                            tt.MoveNext
                                        Loop Until tt.eof
                                    end if

                                    tt.close: set tt = nothing
                                %>
                        </select>   
                    <% end if %>

                    <select  class="no-field" name="cboForma" id="cboForma" onChange="Requery();">
                        <option value="*" <% if Forma = "*"   then response.write " selected" %>>- - Forma - -</option>
                        <%
                            sqlString = "select Forma, Nombre " & _
                                        "from discos_Formas " & _
                                    "where Usuario = '" & usu & "' " 
                            
                            Select Case Tipo 
                                Case "DM", "VM": sqlString = SqlString & "and Musica = 1 "
                                Case "PE": sqlString = SqlString & "and Video = 1 "
                                Case "JU": sqlString = SqlString & "and Juegos = 1 "
                                Case "SO": sqlString = SqlString & "and Software = 1 "
                                Case "LI": sqlString = SqlString & "and Libros = 1 "
                                Case "HW": sqlString = SqlString & "and Hardware = 1 "
                            end Select

                            sqlString = sqlString & "order by Nombre "

                            set tt = con.execute(sqlString)

                            if not (tt.bof or tt.eof) then
                                Do
                                    response.write "<option value='" & tt("Forma") & "'"
                                    if Forma = tt("forma") then response.write " selected" 
                                    response.write ">" & tt("Nombre") & "</option>"

                                    tt.MoveNext
                                Loop Until tt.eof
                            end if

                            tt.close: set tt = nothing
                        %>
                    </select>      

                    <select class="no-field" name="txtAmo" id="txtAmo" onChange="Requery();">
                        <optgroup label="Año de Edición">
                            <option value="1" <% if amo = "1" then response.write " selected" %>><%= aActual %></option>
                            <option value="2" <% if amo = "2" then response.write " selected" %>><%= (aActual - 1) %></option>
                            <option value="3" <% if amo = "3" then response.write " selected" %>>Ultimos 2 Años</option>
                            <option value="4" <% if amo = "4" then response.write " selected" %>>Ultimos 5 Años</option>
                            <option value="5" <% if amo = "5" then response.write " selected" %>>Ultimos 10 Años</option>
                            <option value="6" <% if amo = "6" then response.write " selected" %>>Ultimos 15 Años</option>
                            <option value="7" <% if amo = "7" then response.write " selected" %>>Ultimos 20 Años</option>
                        </optgroup>

                        <optgroup label="Año de Compra">
                            <option value="8"  <% if amo =  "8" then response.write " selected" %>><%= aActual %></option>
                            <option value="9"  <% if amo =  "9" then response.write " selected" %>><%= (aActual - 1) %></option>
                            <option value="10" <% if amo = "10" then response.write " selected" %>>Ultimos 2 Años</option>
                            <option value="11" <% if amo = "11" then response.write " selected" %>>Ultimos 5 Años</option>
                            <option value="12" <% if amo = "12" then response.write " selected" %>>Ultimos 10 Años</option>
                            <option value="13" <% if amo = "13" then response.write " selected" %>>Ultimos 15 Años</option>
                            <option value="14" <% if amo = "14" then response.write " selected" %>>Ultimos 20 Años</option>
                        </optgroup>

                        <optgroup label="- - - - - - - - - - - -">
                            <option value="0" <% if amo = "0" then response.write " selected" %>>Ver Todo</option>
                        </optgroup>
                    </select>  

                    
                    <select class="no-field" name="cboOrden" id="cboOrden" onChange="Requery();">
                        <option value="1" <% if Ordenamiento = 1 then response.write " selected" %>>▲ A&ntilde;o</option>
                        <option value="2" <% if Ordenamiento = 2 then response.write " selected" %>>▲ Titulo</option>
                        <option value="3" <% if Ordenamiento = 3 then response.write " selected" %>>▲ Casa</option>
                        <option value="4" <% if Ordenamiento = 4 then response.write " selected" %>>▲ Medios</option>
                        <option value="5" <% if Ordenamiento = 5 then response.write " selected" %>>▲ Precio</option>

                        <option value="6" <% if Ordenamiento = 6 then response.write " selected" %>>▼ A&ntilde;o</option>
                        <option value="7" <% if Ordenamiento = 7 then response.write " selected" %>>▼ Titulo</option>
                        <option value="8" <% if Ordenamiento = 8 then response.write " selected" %>>▼ Casa</option>
                        <option value="9" <% if Ordenamiento = 9 then response.write " selected" %>>▼ Medios</option>
                        <option value="10" <% if Ordenamiento = 10 then response.write " selected" %>>▼ Precio</option>            
                    </select>           
                </div>
                
                <div style="flex: 0 0 10%; text-align: right;">
                    <button type="button" class="form-btn verde" onClick="nuevo_paquete()">
                        <i class=' fa fa-plus fa-xl' title='Nuevo Objeto'></i>
                    </button>
                </div>
            </div>        

            <div class="main ">
                <div class="tabla-wrapper">
                    <table class="tabla tabla-green">
                        <thead>
                            <tr>
                                <th class="sticky" style="width:  8%; text-align: center;">Portada</th>
                                <th class="sticky" style="width: 45%; text-align: center;">Titulo</th>
                                <th class="sticky" style="width:  5%; text-align: center;">Medios</th>
                                <th class="sticky" style="width:  8%; text-align: center;">Precio</th>
                                <th class="sticky" style="width:  8%; text-align: center;">Forma</th>
                                <th class="sticky" style="width:  6%; text-align: center;">3D</th>
                                <th class="sticky" style="width: 20%; text-align: center;">Acciones</th>
                            </tr>
                        </thead>

                        <tbody>
                            <%
                                if not (t.bof or t.eof) then
                                    response.write "<tbody>"

                                    Do
                                        cuantos = cuantos + 1
                                        objetos = objetos + t("Medios")
                                        total = total + t("Precio")

                                        vinculo = "editar.asp?m=" & t("Paquete")

                                        response.write "<tr>"

                                            response.write "<td style='text-align: center;'>"
                                                response.write "<a class='linea' href='" & vinculo & "'>"
                                                    fotoObjeto = "/perfiles/" & usu & "/medios/" & t("Paquete") & "_s.jpg"
                                                    fotoError = "/perfiles/" & usu & "/discos/foto.jpg"

                                                    %><img src="<%= fotoObjeto %>" onerror="this.src='<%= fotoError %>'" width="60"><%
                                                response.write "</a>"
                                            response.write "</td>"

                                            response.write "<td style='text-align: left;'>"
                                                %><a class="linea" style="font-family: 'Ruda Bold';" href="<%= vinculo %>"><%
                                                    miTipo = ""
                                                    if InStr(t("ListaTipos"),"DM") then miTipo = miTipo & "Musica, "
                                                    if InStr(t("ListaTipos"),"VM") then miTipo = miTipo & "Video Clips / Concierto en Video, "
                                                    if InStr(t("ListaTipos"),"PE") then miTipo = miTipo & "Pelicula o Serie de TV, "
                                                    if InStr(t("ListaTipos"),"JU") then miTipo = miTipo & "Video Juego, "
                                                    if InStr(t("ListaTipos"),"SO") then miTipo = miTipo & "Software o Aplicacion, "
                                                    if InStr(t("ListaTipos"),"LI") then miTipo = miTipo & "Libro o eBook, "
                                                    if InStr(t("ListaTipos"),"HW") then miTipo = miTipo & "Equipo o Periférico, "
                                                    if miTipo <> "" then miTipo = left(miTipo, len(miTipo) - 2)                    

                                                    response.write t("Titulo") & "<br/ >"

                                                    if len(trim(t("ListaInDirAu"))) > 0 then
                                                        response.write "<span style='font-weight: normal; font-style: italic;'>"
                                                            response.write t("ListaInDirAu") & "<br/ >"
                                                        response.write "</span>"
                                                    end if

                                                    response.write "<span style='font-weight: normal;'>"
                                                        response.write t("Amo") & ", " & t("Casa") & "&nbsp;&nbsp;(" & miTipo & ")<br />"
                                                    response.write "</span>"
                                                response.write "</a>"
                                            response.write "</td>"

                                            response.write "<td style='text-align: center;'>"
                                                response.write "<a class='linea' href='" & vinculo & "'>"
                                                    response.write t("Medios")
                                                response.write "</a>"
                                            response.write "</td>"

                                            response.write "<td style='text-align: center;'>"
                                                response.write "<a class='linea' href='" & vinculo & "'>"
                                                    response.write FormatNumber(t("Precio"))
                                                response.write "</a>"
                                            response.write "</td>"                  

                                            response.write "<td style='text-align: center;'>"
                                                response.write "<a class='linea' href='" & vinculo & "'>"
                                                    response.write "<img src='/perfiles/" & usu & "/discos/" & t("Icono_Forma") & "' width='60' />"
                                                response.write "</a>"
                                            response.write "</td>"

                                            response.write "<td style='text-align: center;'>"
                                                response.write "<a class='linea' href='" & vinculo & "'>"
                                                    if t("Medio3D") = 1 then
                                                        response.write "<img src='/perfiles/" & usu & "/discos/icono_3D.gif' width='60' />"
                                                    else
                                                        response.write "&nbsp;"
                                                    end if
                                                response.write "</a>"
                                            response.write "</td>"

                                            response.write "<td style='text-align: center; padding:0px; width:12%' class='borde'>"
                                                %><button type="button"  class="form-btn azul" onClick="copiarPaquete('<%= t("Paquete") %>')"><%
                                                    response.write "<i class=' fa fa-copy fa-xl' title='Duplicar Paquete'></i>"
                                                response.write "</button>&nbsp;"

                                                %><button type="button"  class="form-btn rojo" onClick="borrarPaquete('<%= t("Paquete") %>')"><%
                                                    response.write "<i class=' fa fa-trash fa-xl' title='Borrar Paquete'></i>"
                                                response.write "</button>"
                                            response.write "</td>"

                                        response.write "</tr>"

                                        t.MoveNext
                                    Loop Until t.eof

                                    response.write "</tbody>"
                                end if

                                t.close: set t = nothing
                            %>                    
                        </tbody>

                        <tfoot>
                            <tr>
                                <td class="sticky" colspan="7" style="text-align: center;">
                                    <%
                                        Select Case cuantos
                                            case 0: response.write "No hay Paquetes"
                                            case 1: response.write "Un Paquete"
                                            case else
                                                %>
                                                    <div style="width: 100%; display: flex; justify-content: space-between;">
                                                        <span>Paquetes: <%= Cuantos %></span>
                                                        <span>Objetos: <%= objetos %></span>
                                                        <span>Monto: <%= FormatNumber(Total) %></span>
                                                    </div>                            
                                                <%
                                        end Select
                                    %>                            
                                </td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
        </form>
  
        <br />

        <script type="text/javascript">
            function Requery() {
                document.getElementById("formulario").submit();
            }

            function ordenar(campo) {
                var grupo = document.getElementById("cboVerGrupo").value;
                var vinculo = "lista.asp?g=" + grupo + "&o=" + campo;

                window.location.href = vinculo;
            }

            function copiarPaquete(paquete) {
                var confirmacion = confirm("Desea duplicar el paquete seleccionado?");
                var vinculo = "paquete_duplicar.asp?p=" + paquete;

                if (confirmacion) {     
                    window.location.href = vinculo;
                };          
            }

            function borrarPaquete(paquete) {
                var confirmacion = confirm("Desea borrar el paquete seleccionado?");
                var vinculo = "paquete_borrar.asp?p=" + paquete;

                if (confirmacion) {     
                    window.location.href = vinculo;
                };
            }   

            function nuevo_paquete() {
                var vinculo = "nuevo_paquete.asp";
                window.location.href = vinculo;
            }
        </script>   

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
        <% con.close: set con = nothing %> 
    </body>
</html>