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

<%
    Sub ListaFiltro(filtro, tituloFiltro)
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

        lsqlString = "SELECT Usuario, Paquete, Amo, AmoCompra, Titulo, Casa, CasaDisquera, Tienda, Medios, Forma, Precio, Coleccion, VerComo, Medio3D, " & _
                        " ListaFormas, ListaTipos, ListaInDirAu, ListaPlataformas, ListaGeneros, Icono_Forma " & _
                    "FROM discos_MediosPorForma " & _
                    "WHERE (Usuario = '" & lUsuario & "') " & _
                    "AND (" & filtro & ") " & _
                    "ORDER BY Amo DESC, Paquete Desc;"

        set lcon = Server.CreateObject("ADODB.Connection")
        lcon.open Application("Conn")
        set ltt = lcon.Execute(lsqlString)     

        %>
            <table class="tabla tabla-carbon">
                <thead>
                    <tr>
                        <th class="sticky" style="width:  8%; text-align: center;">Portada</th>
                        <th class="sticky" style="width: 50%; text-align: center;">Titulo</th>
                        <th class="sticky" style="width:  5%; text-align: center;">Medios</th>
                        <th class="sticky" style="width:  8%; text-align: center;">Precio</th>
                        <th class="sticky" style="width:  8%; text-align: center;">Forma</th>
                        <th class="sticky" style="width:  7%; text-align: center;">3D</th>
                        <th class="sticky" style="width: 14%; text-align: center;">Acciones</th>
                    </tr>
                </thead>
  
                <%
                    sw = -1

                    if not (ltt.bof or ltt.eof) then
                        response.write "<tbody>"

                        Do
                            cuantos = cuantos + 1
                            objetos = objetos + ltt("Medios")
                            total = total + ltt("Precio")

                            vinculo = "/forma/desktop/apps/discos/inventario/medios/editar.asp?m=" & ltt("Paquete")

                            response.write "<tr>"

                                response.write "<td style='text-align: center;>"
                                    response.write "<a class='linea' href='" & vinculo & "'>"
                                        fotoObjeto = "/perfiles/" & lUsuario & "/medios/" & ltt("Paquete") & "_s.jpg"
                                        fotoError = "/perfiles/" & lUsuario & "/discos/foto.jpg"

                                        %><img src="<%= fotoObjeto %>" onerror="this.src='<%= fotoError %>'" width="60"><%
                                    response.write "</a>"
                                response.write "</td>"

                                miTipo = ""

                                if InStr(ltt("ListaTipos"),"DM") then miTipo = miTipo & "Musica, "
                                if InStr(ltt("ListaTipos"),"VM") then miTipo = miTipo & "Video Clips / Concierto en Video, "
                                if InStr(ltt("ListaTipos"),"PE") then miTipo = miTipo & "Pelicula o Serie de TV, "
                                if InStr(ltt("ListaTipos"),"JU") then miTipo = miTipo & "Video Juego, "
                                if InStr(ltt("ListaTipos"),"SO") then miTipo = miTipo & "Software o Aplicacion, "
                                if InStr(ltt("ListaTipos"),"LI") then miTipo = miTipo & "Libro o eBook, "
                                if InStr(ltt("ListaTipos"),"HW") then miTipo = miTipo & "Equipo o Periférico, "
                                if miTipo <> "" then miTipo = left(miTipo, len(miTipo) - 2)                             

                                response.write "<td style='text-align: left;'>"
                                    response.write "<a class='linea' href='" & vinculo & "'>"
                                        %><span style="font-family: 'Ruda Bold';"><%
                                            response.write ltt("Titulo") & "<br/ >"
                                        response.write "</span>"

                                        if len(trim(ltt("ListaInDirAu"))) > 0 then
                                            response.write "<span class='normal renglon' style='font-weight: normal; font-style: italic;'>"
                                                response.write ltt("ListaInDirAu") & "<br/ >"
                                            response.write "</span>"
                                        end if

                                        response.write "<span class='normal renglon' style='font-weight: normal;'>"
                                            response.write ltt("Amo") & ", " & ltt("Casa") & "&nbsp;&nbsp;(" & miTipo & ")<br />"
                                        response.write "</span>"
                                    response.write "</a>"
                                response.write "</td>"

                                response.write "<td style='text-align: center;'>"
                                    response.write "<a class='linea'  href='" & vinculo & "'>"
                                        response.write ltt("Medios")
                                    response.write "</a>"
                                response.write "</td>"

                                response.write "<td style='text-align: center;'>"
                                    response.write "<a class='linea'  href='" & vinculo & "'>"
                                        response.write FormatNumber(ltt("Precio"))
                                    response.write "</a>"
                                response.write "</td>"                  

                                response.write "<td style='text-align: center;'>"
                                    response.write "<a class='linea'  href='" & vinculo & "'>"
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
                                        response.write "<i class='fa fa-copy fa-xl' title='Duplicar Paquete'></i>"
                                    response.write "</button>&nbsp;"

                                    %><button class="form-btn rojo" onClick="borrarPaquete('<%= ltt("Paquete") %>')"><%
                                        response.write "<i class='fa fa-trash fa-xl' title='Borrar Paquete'></i>"
                                    response.write "</button>"
                                response.write "</td>" 

                            response.write "</tr>"              

                            ltt.MoveNext
                        Loop Until ltt.eof

                        response.write "</tbody>"
                    end if
                %>

                <tfoot>
                    <tr>
                        <td class="sticky" colspan="3" style="text-align: center;">
                            <%
                                Select Case cuantos
                                    case 0: response.write "No hay Paquetes"
                                    case 1: response.write "Un Paquete"
                                    case else
                                        response.write "Paquetes: " & cuantos
                                end Select
                            %>                                
                        </td>

                        <td class="sticky" colspan="2" style="text-align: center;">
                            <%= "Objetos: " & objetos %>
                        </td>

                        <td class="sticky" colspan="2" style="text-align: center;">
                            <%= "Monto: " & FormatNumber(Total) %>
                        </td>                        
                    </tr>
                </tfoot>                
            </table>             
        <%

        ltt.close: set ltt = nothing
        lcon.close: set lcon = nothing
    end Sub
%>     