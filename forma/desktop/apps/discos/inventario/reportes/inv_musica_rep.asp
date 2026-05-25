<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Reporte de Musica por Interpretes </title>        
        
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "discos"
            thisProcess = "discos.0120"
            SysLockOut
        %>    
                
        <style>
            body { overflow: auto; }
            
            table.principal td {
                border-collapse: collapse;
                font-family: Verdana,sans-serif;
                font-size: 12px;
                padding: 5px;
            }

            .celda {
                border: 1px solid rgb(200, 200, 200); 
            }

            .par {
                background-color: rgb(230, 230, 230);
                font-family: Verdana,sans-serif;
                font-size: 12px;
            }

            .impar {               
                background-color: rgb(255, 255, 255);
                font-family: Verdana, sans-serif;
                font-size: 12px;
            }     

            .titulo {
                background-color: rgb(0, 0, 0);
                color: rgb(255, 255, 255);
                font-family: Verdana, sans-serif;
                font-size: 12px;
                text-align: center;
                padding: 5px;
            }   

            .center {
                margin-left: auto;
                margin-right: auto;
            }                
        </style>

        <%
            dim con, t, sqlString, usu, editor, Agrupar, campo, sw, campoRuptura
            dim listaFormas, listaPlataformas, k, ruptura, PrimeraRuptura
            dim medios, mediosRuptura
            dim carpetas, formas, tiendas, ll, fCampo

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")            

            sub Informe()
                %>
                    <table style="width:95%;" class="center principal">
                        <tr>
                            <td style="width:100%; font-size: 18px; text-align:center;">
                                <%
                                    if request.form("chk_ruptura") = 1 then
                                        response.write "INVENTARIO DE MUSICA POR INICIAL DEL INTERPRETE"
                                    else
                                        response.write "INVENTARIO DE MUSICA POR NOMBRE DEL INTERPRETE"
                                    end if                                
                                %>
                            </td>
                        </tr>                                   
                    </table>

                    <table class="titulo center principal" style="width:95%; background-color; rgb(0, 0, 0);">
                        <tr>
                            <td style="width: 20%; font-size: 16px; padding: 10px; text-align: center;">Interprete</td>
                            <td style="width:  5%; font-size: 16px; padding: 10px; text-align: center;">Año</td>
                            <td style="width: 35%; font-size: 16px; padding: 10px; text-align: center;">Titulo</td>
                            <td style="width: 15%; font-size: 16px; padding: 10px; text-align: center;">Genero</td>
                            <td style="width: 15%; font-size: 16px; padding: 10px; text-align: center;">Tienda</td>
                            <td style="width: 10%; font-size: 16px; padding: 10px; text-align: center;">Forma</td>
                        </tr>
                    </table>                
                <%

                '
                ' Leer las formas seleccionadas
                '
                set ll = con.execute("SELECT DISTINCT Forma, NombreForma " & _
                                        "FROM dbo.discos_Rep_Musica_InDirAu " & _
                                        "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                        "ORDER BY NombreForma;")
                
                if not (ll.bof or ll.eof) then
                    formas = "('-*-'"

                    do
                        fCampo = "f" & ll("Forma")

                        if request.form(fCampo) = 1 then
                            formas = formas & ", '" &ll("Forma") & "'"
                        end if

                        ll.movenext
                    loop until ll.eof

                    formas = formas & ")"
                end if                    

                ll.close: set ll = nothing


                '
                ' Leer las carpetas seleccionadas
                '
                set ll = con.execute("SELECT DISTINCT Carpeta, NombreCarpeta " & _
                                        "FROM dbo.discos_Rep_Musica_InDirAu " & _
                                        "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                        "ORDER BY NombreCarpeta;")
                
                if not (ll.bof or ll.eof) then
                    carpetas  = "('-*-'"

                    do
                        fCampo = "c" & ll("Carpeta")

                        if request.form(fCampo) = 1 then
                            carpetas = carpetas & ", '" & ll("Carpeta") & "'"
                        end if

                        ll.movenext
                    loop until ll.eof

                    carpetas = carpetas & ")"
                end if                    

                ll.close: set ll = nothing


                '
                ' Leer las tiendas seleccionadas
                '
                set ll = con.execute("SELECT DISTINCT Tienda, NombreTienda " & _
                                        "FROM dbo.discos_Rep_Musica_InDirAu " & _
                                        "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                        "ORDER BY NombreTienda;")
                
                if not (ll.bof or ll.eof) then
                    tiendas  = "('-*-'"

                    do
                        fCampo = "t" & ll("tienda")

                        if request.form(fCampo) = 1 then
                            tiendas = tiendas & ", '" & ll("tienda") & "'"
                        end if

                        ll.movenext
                    loop until ll.eof

                    tiendas = tiendas & ")"
                end if                    

                ll.close: set ll = nothing


                '
                ' Armamos el Comando SQL...
                '

                sqlString = "SELECT LEFT(InDirAu, 1) AS Inicial, InDirAu AS Interprete, AEdicion, Titulo, NombreTipo AS Genero, NombreForma AS Forma, NombreTienda as Tienda " & _
                                "FROM discos_Rep_Musica_InDirAu " & _
                                "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " 
                
                if carpetas <> "" then
                    sqlString = sqlString & "AND (Carpeta IN " & carpetas & ") "
                end if

                if formas <> "" then
                    sqlString = sqlString & "AND (Forma IN " & formas & ") "
                end if

                if tiendas <> "" then
                    sqlString = sqlString & "AND (tienda IN " & tiendas & ") "
                end if

                sqlString = sqlString & "ORDER BY Inicial, Interprete, AEdicion;"


                '
                ' Procesamos el Informe
                '

                set t = con.execute(sqlString)
            
                if not (t.bof or t.eof) then
                    sw = -1
                    medios = 0
                    mediosRuptura = 0
                    ruptura = "/*"

                    %>
                        <table style="width:95%;" class="center principal">
                            <%  
                                Do
                                    if request.form("chk_ruptura") = 1 then
                                        chk_ruptura = t("Inicial")
                                    else
                                        chk_ruptura = t("Interprete")
                                    end if

                                    if chk_ruptura <> ruptura then
                                        if ruptura <> "/*" then
                                            %>
                                                <tr style = "background-color: rgb(135, 135, 135); font-family: Arial; color: rgb(255, 255, 255);">
                                                    <td colspan="6" style = "width: 100%; padding: 10px; text-align: right; font-size: 14px;">
                                                        <% 
                                                            response.write "Objetos:&nbsp;&nbsp;" & mediosRuptura & "&nbsp;&nbsp;<br />"
                                                            medios = medios + mediosRuptura
                                                            mediosRuptura = 0
                                                        %>
                                                    </td>                                               
                                                </tr>

                                                <tr style = "border-style: none; background-color: transparent;">
                                                    <td colspan="6">
                                                        &nbsp;
                                                    </td>
                                                </tr>
                                            <%                                        
                                        end if

                                        ruptura = chk_ruptura

                                        %>
                                            <tr style = "background-color: rgb(90, 90, 90); font-wright: bold; font-family: Arial; color: rgb(255, 255, 255);">
                                                <td colspan="6" style = "width: 100%; padding: 10px; font-size: 14px;">
                                                    <%= chk_ruptura %>
                                                </td>
                                            </tr>                                        
                                        <%
                                    end if

                                    sw = -1 * sw

                                    %>
                                        <tr class="<% if sw = 1 then response.write "par" else response.write "impar"%>">
                                            <td class="celda" style="width: 20%; text-align: left;"><%= t("Interprete") %></td>
                                            <td class="celda" style="width:  5%; text-align: center;"><%= t("AEdicion") %></td>
                                            <td class="celda" style="width: 35%; text-align: left;"><%= t("Titulo") %></td>
                                            <td class="celda" style="width: 15%; text-align: left;"><%= t("Genero") %></td>
                                            <td class="celda" style="width: 15%; text-align: left;"><%= t("Tienda") %></td>
                                            <td class="celda" style="width: 10%; text-align: left;"><%= t("Forma") %></td>
                                        </tr>                                        
                                    <%

                                    mediosRuptura = mediosRuptura + 1

                                    t.MoveNext
                                Loop until (t.eof)
                            %>

                            <tr style = "background-color: rgb(135, 135, 135); font-family: Arial; color: rgb(255, 255, 255);">
                                <td colspan="6" style = "width: 100%; padding: 10px; text-align: right; font-size: 14px;">
                                    <% 
                                        response.write "Objetos:&nbsp;&nbsp;" & mediosRuptura & "&nbsp;&nbsp;<br />"
                                        medios = medios + mediosRuptura
                                    %>
                                </td>                                               
                            </tr>
                        </table>

                        <table style="width:95%;" class="titulo center principal">
                            <tr><td style="text-align: right; font-size: 14px;">Total de Objetos:&nbsp;&nbsp;<%= Medios %>&nbsp;&nbsp;</td></tr>
                        </table>

                        <br />    
                    <%
                end if

                t.close: set t = nothing                      
            end sub
        %>
    </head>

    <body style="background-color: rgb(255,255,255);">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->      
        <br />
        <%  
            Informe    
            con.close: set con = nothing                    
        %>   
    </body>
</html>