<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<%
    function CarpetaPorDefecto()
        dim cc, tt

        sqlString = "SELECT Codigo FROM discos_Carpetas WHERE Usuario = '" & Request.QueryString("Usuario") & "' AND PorDefecto = 1;"

        set cc = Server.CreateObject("ADODB.Connection")
        cc.open Application("Conn")
        set tt = cc.execute(sqlString)

            CarpetaPorDefecto = tt("Codigo")

        tt.close: set tt = nothing
        cc.close: set cc = nothing
    end function         


    '
    ' Abrimos la tabla y llenamos los datos
    '

    dim con, t, tt, sqlString, usu, vinculo, sw
    dim cuantos, objetos, total, aActual, lsqlString
    dim tipo, forma, amo, folder, ordenamiento
    
    usu = Request.Cookies("Usuario")

    folder = Request.QueryString("folder") 
    tipo = Request.QueryString("tipo") 
    forma = Request.QueryString("forma") 
    plataforma = Request.QueryString("plataforma") 
    amo = Request.QueryString("amo") 
    ordenamiento = Request.QueryString("orden")

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

    %>
        <table class="tabla tabla-green">
            <thead>
                <tr>
                    <th class="sticky" style="width:  8%; text-align: center;">Portada</th>
                    <th class="sticky" style="width: 65%; text-align: center;">Titulo</th>
                    <th class="sticky" style="width:  5%; text-align: center;">Medios</th>
                    <th class="sticky" style="width:  8%; text-align: center;">Precio</th>
                    <th class="sticky" style="width:  8%; text-align: center;">Forma</th>
                    <th class="sticky" style="width:  6%; text-align: center;">3D</th>
                </tr>
            </thead>

            <tbody>    
                <%
                    set t = con.Execute(sqlString)        
                        if not (t.bof or t.eof) then
                            Do
                                cuantos = cuantos + 1
                                objetos = objetos + t("Medios")
                                total = total + t("Precio")

                                vinculo = "loadInWindow('discos_lista', '" & "/forma/snippets/recursos/discos_ver.asp?paquete=" & t("Paquete") & "&parent=discos_lista')"

                                fotoObjeto = "/perfiles/" & Request.Cookies("Usuario") & "/medios/" & t("Paquete") & "_s.jpg"
                                fotoError = "/perfiles/" & Request.Cookies("Usuario") & "/discos/foto.jpg"

                                %>
                                    <tr>
                                        <td style='text-align: center;'>
                                            <a class='discos_lista_linea' onclick="<%= vinculo %>" >
                                                <img src="<%= fotoObjeto %>" onerror="this.src='<%= fotoError %>'" width="60">
                                            </a>
                                        </td>

                                        <td style='text-align: left;'>
                                            <a class="discos_lista_linea" style="font-family: 'Ruda Bold';" onclick="<%= vinculo %>" >
                                                <%
                                                    miTipo = ""
                                                    
                                                    if InStr (t("ListaTipos"),"DM") then miTipo = miTipo & "Musica, "
                                                    if InStr (t("ListaTipos"),"VM") then miTipo = miTipo & "Video Clips / Concierto en Video, "
                                                    if InStr (t("ListaTipos"),"PE") then miTipo = miTipo & "Pelicula o Serie de TV, "
                                                    if InStr (t("ListaTipos"),"JU") then miTipo = miTipo & "Video Juego, "
                                                    if InStr (t("ListaTipos"),"SO") then miTipo = miTipo & "Software o Aplicacion, "
                                                    if InStr (t("ListaTipos"),"LI") then miTipo = miTipo & "Libro o eBook, "
                                                    if InStr (t("ListaTipos"),"HW") then miTipo = miTipo & "Equipo o Periférico, "

                                                    if miTipo <> "" then miTipo = left(miTipo, len(miTipo) - 2)  
                                                            
                                                    response.write t("Titulo") & "<br />"

                                                    if len(trim (t("ListaInDirAu"))) > 0 then
                                                        response.write "<span style='font-weight: normal; font-style: italic;'>"
                                                            response.write t("ListaInDirAu") & "<br />"
                                                        response.write "</span>"
                                                    end if

                                                    response.write "<span style='font-weight: normal;'>"
                                                        response.write t("Amo") & ", " & t("Casa") & "&nbsp;&nbsp;(" & miTipo & ")<br />"
                                                    response.write "</span>"
                                                %>
                                            </a>
                                        </td>

                                        <td style='text-align: center;'>
                                            <a class='discos_lista_linea' onclick="<%= vinculo %>" >
                                                <%= t("Medios") %>
                                            </a>
                                        </td>

                                        <td style='text-align: center;'>
                                            <a class='discos_lista_linea' onclick="<%= vinculo %>" >
                                                <%= FormatNumber(t("Precio")) %>
                                            </a>
                                        </td>

                                        <td style='text-align: center;'>
                                            <a class='discos_lista_linea' onclick="<%= vinculo %>" >
                                                <img src='/perfiles/<%= Request.Cookies("Usuario") %>/discos/<%= t("Icono_Forma") %>' width='60' />
                                            </a>
                                        </td>

                                        <td style='text-align: center;'>
                                            <a class='discos_lista_linea' onclick="<%= vinculo %>" >
                                                <%
                                                    if t("Medio3D") = 1 then
                                                        response.write "<img src='/perfiles/" & Request.Cookies("Usuario") & "/discos/icono_3D.gif' width='60' />"
                                                    else
                                                        response.write "&nbsp;"
                                                    end if
                                                %>
                                            </a>
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
                    <td class="sticky" colspan="6" style="text-align: center;">
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
    <% 
    
    con.close: set con = nothing 
%>