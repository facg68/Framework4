<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>

        <title>Buscar Medios</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "discos"
            thisProcess = "discos.0120"
            SysLockOut

            Sub ListaFiltro(filtro, tituloFiltro, Ordenamiento, Editor)
                dim lcon, ltt, lsqlString, lUsuario, lSW
                dim lcuantos, lobjetos, ltotal

                lUsuario = Request.Cookies("Usuario")
                lcuantos = 0
                lobjetos = 0
                ltotal = 0.00

                '
                ' Creamos la cadena de conexión, dependiendo de los
                ' datos del filtro, o generamos una cadena nueva
                '
                Select Case Editor
                    Case "VM"
                        lsqlString = "SELECT Usuario, Paquete, Objeto, Amo, AmoCompra, Titulo, Casa, CasaDisquera, Tienda, Medios, Forma, Precio, Coleccion, VerComo, Medio3D, " & _
                                           " ListaFormas, ListaTipos, ListaInDirAu, ListaPlataformas, ListaGeneros, Icono_Forma " & _
                                       "FROM discos_MediosPorForma_Filtro_VM " & _
                                      "WHERE (Usuario = '" & lUsuario & "') " & _
                                        "AND (" & filtro & ") " & _
                                   "ORDER BY "                        
                    Case Else
                        lsqlString = "SELECT Usuario, Paquete, Objeto, Amo, AmoCompra, Titulo, Casa, CasaDisquera, Tienda, Medios, Forma, Precio, Coleccion, VerComo, Medio3D, " & _
                                           " ListaFormas, ListaTipos, ListaInDirAu, ListaPlataformas, ListaGeneros, Icono_Forma " & _
                                       "FROM discos_MediosPorForma_Filtro " & _
                                      "WHERE (Usuario = '" & lUsuario & "') " & _
                                        "AND (" & filtro & ") " & _
                                        "AND (ListaTipos LIKE '%" & Editor & "%')" & _
                                  "ORDER BY "                    
                End Select                

                select case ordenamiento
                    case 1: lsqlString = lsqlString & "Amo;"
                    case 2: lsqlString = lsqlString & "Titulo;"
                    case 3: lsqlString = lsqlString & "Amo DESC;"
                    case 4: lsqlString = lsqlString & "Titulo DESC;"
                end select                      

                set lcon = Server.CreateObject("ADODB.Connection")
                lcon.open Application("Conn")
                set ltt = lcon.Execute(lsqlString)     

                %>
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
                                    sw = -1

                                    if not (ltt.bof or ltt.eof) then
                                        response.write "<tbody>"

                                        Do
                                            sw = -1 * cint(sw)              
                                            cuantos = cuantos + 1
                                            objetos = objetos + ltt("Medios")
                                            total = total + ltt("Precio")

                                            vinculo = "/forma/desktop/apps/discos/inventario/medios/editar.asp?m=" & ltt("Paquete")

                                            response.write "<tr>"

                                                response.write "<td style='text-align: center;'>"
                                                    response.write "<a class='linea' href='" & vinculo & "'>"
                                                        fotoPaquete = "/perfiles/" & lUsuario & "/medios/" & ltt("Paquete") & "_s.jpg"
                                                        fotoObjeto = "/perfiles/" & lUsuario & "/medios/" & ltt("Objeto") & "_s.jpg"

                                                        %><img src="<%= fotoObjeto %>" onerror="this.src='<%= fotoPaquete %>'" width="60"><%
                                                    response.write "</a>"
                                                response.write "</td>"

                                                response.write "<td style='text-align: left;'>"
                                                    response.write "<a class='linea' href='" & vinculo & "'>"
                                                        miTipo = ""
                                                        if InStr(ltt("ListaTipos"),"DM") then miTipo = miTipo & "Musica, "
                                                        if InStr(ltt("ListaTipos"),"VM") then miTipo = miTipo & "Video Clips / Concierto en Video, "
                                                        if InStr(ltt("ListaTipos"),"PE") then miTipo = miTipo & "Pelicula o Serie de TV, "
                                                        if InStr(ltt("ListaTipos"),"JU") then miTipo = miTipo & "Video Juego, "
                                                        if InStr(ltt("ListaTipos"),"SO") then miTipo = miTipo & "Software o Aplicacion, "
                                                        if InStr(ltt("ListaTipos"),"LI") then miTipo = miTipo & "Libro o eBook, "
                                                        if miTipo<> "" then miTipo = left(miTipo, len(miTipo) - 2)                    

                                                        response.write ltt("Titulo") & "<br/ >"
                                                        if len(trim(ltt("ListaInDirAu"))) > 0 then
                                                            response.write "<span style='font-weight: normal; font-style: italic;'>"
                                                                response.write ltt("ListaInDirAu") & "<br/ >"
                                                            response.write "</span>"
                                                        end if

                                                        response.write "<span style='font-weight: normal;'>"
                                                        response.write ltt("Amo") & ", " & ltt("Casa") & "&nbsp;&nbsp;(" & miTipo & ")<br />"
                                                        response.write "</span>"
                                                    response.write "</a>"
                                                response.write "</td>"

                                                response.write "<td style='text-align: center;'>"
                                                    response.write "<a class='linea' href='" & vinculo & "'>"
                                                        response.write ltt("Medios")
                                                    response.write "</a>"
                                                response.write "</td>"

                                                response.write "<td style='text-align: center;'>"
                                                    response.write "<a class='linea' href='" & vinculo & "'>"
                                                        response.write FormatNumber(ltt("Precio"))
                                                    response.write "</a>"
                                                response.write "</td>"                  

                                                response.write "<td style='text-align: center;'>"
                                                    response.write "<a class='linea' href='" & vinculo & "'>"
                                                        response.write "<img src='/perfiles/" & lUsuario & "/discos/" & ltt("Icono_Forma") & "' width='60' />"
                                                    response.write "</a>"
                                                response.write "</td>"

                                                response.write "<td style='text-align: center;'>"
                                                    response.write "<a class='linea' href='" & vinculo & "'>"
                                                        if ltt("Medio3D") = 1 then
                                                            response.write "<img src='/perfiles/" & lUsuario & "/discos/icono_3D.gif' width='60' />"
                                                        else
                                                            response.write "&nbsp;"
                                                        end if
                                                    response.write "</a>"
                                                response.write "</td>"

                                                response.write "<td style='text-align: center;'>"
                                                    %><button class="form-btn azul" onClick="copiarPaquete('<%= ltt("Paquete") %>')"><%
                                                        response.write "<i class=' fa fa-copy fa-xl' title='Duplicar Paquete'></i>"
                                                    response.write "</button>&nbsp;"

                                                    %><button class="form-btn rojo" onClick="borrarPaquete('<%= ltt("Paquete") %>')"><%
                                                        response.write "<i class=' fa fa-trash fa-xl' title='Borrar Paquete'></i>"
                                                    response.write "</button>"
                                                response.write "</td>"

                                            ltt.MoveNext
                                        Loop Until ltt.eof

                                        response.write "</tbody>"
                                    end if
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
                <%

                ltt.close: set ltt = nothing
                lcon.close: set lcon = nothing
            end Sub    

            Function TituloInDirAu(Editor)
                dim ttCon, tt, sqlCommand

                sqlCommand = "SELECT InDirAu FROM discos_Objetos_Clases WHERE Codigo = '" & Editor & "';"

                set ttCon = Server.CreateObject("ADODB.Connection")
                ttCon.open Application("Conn")
                    set tt = ttCon.Execute(sqlCommand)
                        if not (tt.bof or tt.eof) then
                            TituloInDirAu = tt("InDirAu")
                        else
                            TituloInDirAu = ""
                        end if
                    tt.close: set tt = nothing
                ttCon.Close: set ttCon = nothing
            End Function
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

    <body style="background-color: rgb(235, 235, 235);" onload="InicializarTabla(98);">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    
        <%  
            '
            ' Abrimos la tabla y llenamos los datos
            '
            dim con, t, sqlString, vinculo, sw
            dim InDirAu, Editor, Orden, Cuantos
            dim tt, Usuario

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            cuantos = 0
            Usuario = Request.Cookies("Usuario")

            InDirAu = Request.Form("cboInDirAu")
            Editor = Request.Form("cboEditor")
            Orden = Request.Form("cboOrden")

            If Orden = "" then Orden = 3
            if Editor = "" then Editor = "DM"   
        %>

        <br />

        <form id="formulario" name="formulario" method="post" action="buscar_medios.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 10%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Buscar
                </div>

                <div style="flex: 0 0 90%; font-family: Ruda; font-size: 16px; text-align: right;">
                    <select class="no-field" name="cboEditor" id="cboEditor" onChange="Requery();">
                        <%
                            sqlString = "SELECT Codigo, Nombre " & _
                                            "FROM  discos_Objetos_Clases " & _
                                        "ORDER BY Nombre;"

                            set tt = con.Execute(sqlString)

                            if not (tt.bof or tt.eof) then
                                Do
                                    response.write "<option value='" & tt("Codigo") & "'"
                                        if tt("Codigo") = Editor then 
                                            response.write " selected" 
                                        end if
                                    response.write ">" & tt("Nombre") & "</option>"
                                    tt.MoveNext
                                Loop Until tt.eof
                            end if

                            tt.close: set tt = nothing
                        %>
                    </select>

                    <% if Editor <> "" then 
                            select case Editor
                                Case "PE" 
                                    sqlString = "SELECT DISTINCT TOP (100) PERCENT Campo " & _
                                                "FROM (SELECT InDirAu AS Campo " & _ 
                                                        "FROM dbo.discos_Objetos AS o " & _
                                                        "WHERE (Editor = 'PE') " & _
                                                        "AND (o.Usuario = '" & Usuario & "')" & _
                                                "UNION SELECT o1p.Protagonista AS Campo " & _
                                                        "FROM dbo.discos_Objetos AS o1 " & _
                                                    "INNER JOIN dbo.discos_Objetos_Protagonistas AS o1p " & _
                                                            "ON o1.Usuario = o1p.Usuario " & _
                                                        "AND o1.Paquete = o1p.Paquete " & _
                                                        "AND o1.Objeto = o1p.Objeto " & _
                                                        "WHERE (o1.Editor = 'PE')" & _
                                                        "AND (o1.Usuario = '" & Usuario & "')" & _
                                                        ") AS q " & _
                                                "WHERE (Campo <> '-') AND (Campo <> '.') AND (Campo <> '') AND (Campo IS NOT NULL);"                         
                                Case "HW"
                                    sqlString = "SELECT DISTINCT c.Nombre AS Campo " & _
                                                    "FROM discos_Objetos AS o " & _ 
                                            "INNER JOIN discos_Paquetes AS p " & _
                                                    "ON o.Usuario = p.Usuario " & _
                                                    "AND o.Paquete = p.Paquete " & _
                                            "INNER JOIN discos_Casas AS c " & _
                                                    "ON p.Usuario = c.Usuario " & _
                                                    "AND p.Casa = c.Codigo " & _
                                                    "WHERE (o.Editor = 'HW') " & _
                                                    "AND (o.Usuario = '" & Usuario & "') " & _
                                                "ORDER BY c.Nombre;" 
                                Case Else
                                    sqlString = "SELECT DISTINCT o.InDirAu AS Campo " & _
                                                "FROM dbo.discos_Objetos AS o " & _
                                                "WHERE (o.Editor = '" & Editor & "') " & _
                                                "AND (o.Usuario = '" & Usuario & "') " & _
                                                "AND (o.InDirAu <> '-') AND (o.InDirAu <> '.') AND (o.InDirAu <> '') AND (o.InDirAu IS NOT NULL) " & _
                                            "ORDER BY o.InDirAu;"  
                            end select 

                            response.write "&nbsp;" & TituloInDirAu(Editor) & "&nbsp;"
                        %>

                        <select class="no-field" name="cboInDirAu" id="cboInDirAu" onChange="Requery();">
                            <option value="" <% if InDirAu = "" then response.write " selected" %>>&nbsp;</option>

                            <%
                                set tt = con.Execute(sqlString)

                                if not (tt.bof or tt.eof) then
                                    Do
                                        response.write "<option value='" & tt("Campo") & "'"
                                            if InDirAu = tt("Campo") then 
                                                response.write " selected" 
                                            end if
                                        response.write ">" & tt("Campo") & "</option>"
                                        tt.MoveNext
                                    Loop Until tt.eof
                                end if
                            %>                          
                        </select>

                        &nbsp;

                        <select class="no-field" name="cboOrden" id="cboOrden" onChange="Requery();">
                            <option value="1" <% if Orden = "1" then response.write " selected" %>>▲ Año</option>
                            <option value="2" <% if Orden = "2" then response.write " selected" %>>▲ Nombre</option>
                            <option value="3" <% if Orden = "3" then response.write " selected" %>>▼ Año</option>
                            <option value="4" <% if Orden = "4" then response.write " selected" %>>▼ Nombre</option>
                        </select>    
                    <% end if %>

                    &nbsp;

                    <input class="no-field" id="ordenamiento" name="ordenamiento" type="hidden" value="<%= Orden %>">                                        
                </div>                
            </div>  

                
            <div class="main" style="background-color: rgba(250, 250, 250, 1);">
                <% If InDirAu <> "" then %>
                    <div class="line">
                        <%
                            ListaFiltro "(ListaInDirAu = '" & InDirAu & "')", "Medios de " & InDirAu, Orden, Editor
                        %>
                    </div>
                <% end if %>
            </div>
        </form>

        <br />   

        <script type="text/javascript">
            wReserva += 5;

            function Requery() {
                document.getElementById("formulario").submit();
            } 
        </script>   

        <!-- #include virtual = "/core/includes/kernel/close.inc" --> 
        <% con.close: set con = nothing %>         
    </body>