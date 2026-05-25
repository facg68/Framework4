<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Informe de Medios</title>     
        
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "discos"
            thisProcess = "discos.0120"
            SysLockOut
        %>    
        
        <style>
            body { overflow: auto; }        

            td.lista {
                font-family: 'Ruda', sans-serif;   
                font-size: 12px;
                padding: 5px;
            }

            .celda {
                border: 1px solid rgb(200, 200, 200); 
            }

            .par {
                background-color: rgb(230, 230, 230);
                font-family: 'Ruda', sans-serif;   
                font-size: 12px;
            }

            .impar {               
                background-color: rgb(255, 255, 255);
                font-family: 'Ruda', sans-serif;   
                font-size: 12px;
            }     

            .titulo {
                background-color: rgb(0, 0, 0);
                color: rgb(255, 255, 255);
                font-family: 'Ruda', sans-serif;   
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
            dim con, t, sqlString, usu, editor, Agrupar, campo, sw, saldo
            dim listaFormas, listaPlataformas, k, ruptura, medios           
            dim saldoRuptura, mediosRuptura, PrimeraRuptura

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            sub Informe_Normal()
                set t = con.execute(sqlString)

                if not (t.bof or t.eof) then
                    sw = -1
                    Saldo = 0
                    medios = 0

                    %>
                        <table style="width:95%;" class="center">
                            <%  
                                Do 
                                    sw = -1 * sw
                                    Saldo = Saldo + t("Precio")
                                    medios = medios + 1
                                %>
                                    <tr class="<% if sw = 1 then response.write "par" else response.write "impar"%>">
                                        <td class="lista celda" style="width:  5%; text-align: center;"><%= t("AEdicion") %></td>
                                        <td class="lista celda" style="width: 30%; text-align: left;"><%
                                            if (InStr(t("Categoria"), "DM") > 0) OR (InStr(t("Categoria"), "VM") > 0) then
                                                response.write t("Interprete") & ": " & t("Titulo") 
                                            else
                                                response.write t("Titulo") 
                                            end if
                                        %></td>
                                        <td class="lista celda" style="width:  8%; text-align: right;"><%= FormatNumber(t("Precio")) %></td>
                                        <td class="lista celda" style="width: 16%; text-align: left;"><%= t("Casa") %></td>
                                        <td class="lista celda" style="width: 12%; text-align: left;"><%
                                            if len(t("Categoria")) > 2 then
                                                response.write "Varios"
                                            else
                                                select case t("Categoria")
                                                    case "DM": response.write "Musica"
                                                    case "PE": response.write "Pelicula"
                                                    case "JU": response.write "Video Juego"
                                                    case "SO": response.write "Software"
                                                    case "LI": response.write "Libro"
                                                    case "VM": response.write "Video Musical"
                                                end select
                                            end if
                                            %></td>
                                        <td class="lista celda" style="width: 12%; text-align: left;"><%= t("Forma") %></td>
                                        <td class="lista celda" style="width: 12%; text-align: left;"><%
                                            if t("NombrePlataforma") = "- No Aplica -" then
                                                response.write "&nbsp;"
                                            else
                                                response.write t("NombrePlataforma")
                                            end if
                                        %></td>
                                        <td class="lista celda" style="width: 5%; text-align: center;"><%
                                            if t("Es_3D") = "3D" then
                                                response.write "3D"
                                            else
                                                response.write "&nbsp;"
                                            end if
                                        %></td>
                                    </tr>                                        
                                <%

                                    t.MoveNext
                                Loop until (t.eof)
                            %>
                        </table>

                        <table style="width:95%;" class="center" >
                            <tr>
                                <td class="lista" style="color: white; background-color: black; font-size: 14px; text-align: right;">
                                    Total de Objetos:&nbsp;&nbsp;<%= Medios %>&nbsp;<br/>
                                    Saldo:&nbsp;&nbsp;<%= FormatNumber(Saldo) %>&nbsp;
                                </td>
                            </tr>
                        </table>  
                    <%
                end if

                t.close: set t = nothing                     
            end sub

            sub Ruptura_Plataforma()
                set t = con.execute(sqlString)

                if not (t.bof or t.eof) then
                    sw = -1
                    Saldo = 0
                    medios = 0
                    ruptura = "/*"

                    %>
                        <table style="width:95%;" class="center">
                            <%  
                                Do
                                    if t("NombrePlataforma") <> ruptura then
                                        if ruptura <> "/*" then
                                            '
                                            ' No es el inicio de la primera ruptura...
                                            ' Presentamos los totales de la ruptura y
                                            ' reiniciamos los contadores
                                            '
                                            %>
                                                <tr style = "background-color: rgb(135, 135, 135); font-family: Arial; color: rgb(255, 255, 255);">
                                                    <td class="lista"  colspan="8" style = "width: 100%; padding: 10px; text-align: right; font-size: 14px;">
                                                        <% 
                                                            response.write "Total de Objetos:&nbsp;&nbsp;" & mediosRuptura & "&nbsp;&nbsp;<br />"
                                                            response.write "Saldo:&nbsp;&nbsp;" & FormatNumber(saldoRuptura) & "&nbsp;&nbsp;"

                                                            saldo = saldo + saldoRuptura
                                                            medios = medios + mediosRuptura
                                                        %>
                                                    </td>                                               
                                                </tr>

                                                <tr style = "border-style: none; background-color: transparent;">
                                                    <td class="lista"  colspan="8">
                                                        &nbsp;
                                                    </td>
                                                </tr>
                                            <%
                                        end if

                                        ruptura = t("NombrePlataforma")
                                        saldoRuptura = 0
                                        mediosRuptura = 0                                        

                                        %>
                                            <tr style = "background-color: rgb(90, 90, 90); font-wright: bold; font-family: Arial; color: rgb(255, 255, 255);">
                                                <td class="lista"  colspan="8" style = "width: 100%; padding: 10px; font-size: 14px;">
                                                    <%= t("NombrePlataforma") %>
                                                </td>
                                            </tr>                                        
                                        <%
                                    end if

                                    sw = -1 * sw
                                    saldoRuptura = saldoRuptura + t("Precio")
                                    mediosRuptura = mediosRuptura + 1
                                %>
                                    <tr class="<% if sw = 1 then response.write "par" else response.write "impar"%>">
                                        <td class="lista celda" style="width:  5%; text-align: center;"><%= t("AEdicion") %></td>
                                        <td class="lista celda" style="width: 30%; text-align: left;"><%
                                            if (InStr(t("Categoria"), "DM") > 0) OR (InStr(t("Categoria"), "VM") > 0) then
                                                response.write t("Interprete") & ": " & t("Titulo") 
                                            else
                                                response.write t("Titulo") 
                                            end if
                                        %></td>
                                        <td class="lista celda" style="width:  8%; text-align: right;"><%= FormatNumber(t("Precio")) %></td>
                                        <td class="lista celda" style="width: 16%; text-align: left;"><%= t("Casa") %></td>
                                        <td class="lista celda" style="width: 12%; text-align: left;"><%
                                            if len(t("Categoria")) > 2 then
                                                response.write "Varios"
                                            else
                                                select case t("Categoria")
                                                    case "DM": response.write "Musica"
                                                    case "PE": response.write "Pelicula"
                                                    case "JU": response.write "Video Juego"
                                                    case "SO": response.write "Software"
                                                    case "LI": response.write "Libro"
                                                    case "VM": response.write "Video Musical"
                                                end select
                                            end if
                                            %></td>
                                        <td class="lista celda" style="width: 12%; text-align: left;"><%= t("Forma") %></td>
                                        <td class="lista celda" style="width: 12%; text-align: left;"><%
                                            if t("NombrePlataforma") = "- No Aplica -" then
                                                response.write "&nbsp;"
                                            else
                                                response.write t("NombrePlataforma")
                                            end if
                                        %></td>
                                        <td class="lista celda" style="width: 5%; text-align: center;"><%
                                            if t("Es_3D") = "3D" then
                                                response.write "3D"
                                            else
                                                response.write "&nbsp;"
                                            end if
                                        %></td>
                                    </tr>                                        
                                <%

                                    t.MoveNext
                                Loop until (t.eof)

                                '
                                ' Se acabaron los datos....
                                ' Presentamos los valores de la ultima ruptura
                                ' y luego el total global
                                '
                                %>
                                    <tr style = "background-color: rgb(135, 135, 135); font-family: Arial; color: rgb(255, 255, 255);">
                                        <td class="lista"  colspan="8" style = "width: 100%; padding: 10px; text-align: right; font-size: 14px;">
                                            <% 
                                                response.write "Total de Objetos:&nbsp;&nbsp;" & mediosRuptura & "&nbsp;&nbsp;<br />"
                                                response.write "Saldo:&nbsp;&nbsp;" & FormatNumber(saldoRuptura) & "&nbsp;&nbsp;"

                                                saldo = saldo + saldoRuptura
                                                medios = medios + mediosRuptura
                                            %>
                                        </td>                                               
                                    </tr>
                                <%
                            %>
                        </table>

                        <table style="width:95%;" class="center" >
                            <tr>
                                <td class="lista" style="color: white; background-color: black; font-size: 14px; text-align: right;">
                                    Total de Objetos:&nbsp;&nbsp;<%= Medios %>&nbsp;<br/>
                                    Saldo:&nbsp;&nbsp;<%= FormatNumber(Saldo) %>&nbsp;<
                                </td>
                            </tr>
                        </table>                        
                    <%
                end if

                t.close: set t = nothing              
            end sub

            sub Ruptura_Metadata()
                set t = con.execute(sqlString)

                if not (t.bof or t.eof) then
                    sw = -1
                    Saldo = 0
                    medios = 0
                    ruptura = "/*"

                    %>
                        <table style="width:95%;" class="center">
                            <%  
                                Do
                                    if t("MetaData") <> ruptura then
                                        if ruptura <> "/*" then
                                            '
                                            ' No es el inicio de la primera ruptura...
                                            ' Presentamos los totales de la ruptura y
                                            ' reiniciamos los contadores
                                            '
                                            %>
                                                <tr style = "background-color: rgb(135, 135, 135); font-family: Arial; color: rgb(255, 255, 255);">
                                                    <td class="lista"  colspan="8" style = "width: 100%; padding: 10px; text-align: right; font-size: 14px;">
                                                        <% 
                                                            response.write "Total de Objetos:&nbsp;&nbsp;" & mediosRuptura & "&nbsp;&nbsp;<br />"
                                                            response.write "Saldo:&nbsp;&nbsp;" & FormatNumber(saldoRuptura) & "&nbsp;&nbsp;"

                                                            saldo = saldo + saldoRuptura
                                                            medios = medios + mediosRuptura
                                                        %>
                                                    </td>                                               
                                                </tr>

                                                <tr style = "border-style: none; background-color: transparent;">
                                                    <td class="lista"  colspan="8">
                                                        &nbsp;
                                                    </td>
                                                </tr>
                                            <%
                                        end if

                                        ruptura = t("MetaData")
                                        saldoRuptura = 0
                                        mediosRuptura = 0                                        

                                        %>
                                            <tr style = "background-color: rgb(90, 90, 90); font-wright: bold; font-family: Arial; color: rgb(255, 255, 255);">
                                                <td class="lista"  colspan="8" style = "width: 100%; padding: 10px; font-size: 14px;">
                                                    <%= t("MetaData") %>
                                                </td>
                                            </tr>                                        
                                        <%
                                    end if

                                    sw = -1 * sw
                                    saldoRuptura = saldoRuptura + t("Precio")
                                    mediosRuptura = mediosRuptura + 1
                                %>
                                    <tr class="<% if sw = 1 then response.write "par" else response.write "impar"%>">
                                        <td class="lista celda" style="width:  5%; text-align: center;"><%= t("AEdicion") %></td>
                                        <td class="lista celda" style="width: 30%; text-align: left;"><%
                                            if (InStr(t("Categoria"), "DM") > 0) OR (InStr(t("Categoria"), "VM") > 0) then
                                                response.write t("Interprete") & ": " & t("Titulo") 
                                            else
                                                response.write t("Titulo") 
                                            end if
                                        %></td>
                                        <td class="lista celda" style="width:  8%; text-align: right;"><%= FormatNumber(t("Precio")) %></td>
                                        <td class="lista celda" style="width: 16%; text-align: left;"><%= t("Casa") %></td>
                                        <td class="lista celda" style="width: 12%; text-align: left;"><%
                                            if len(t("Categoria")) > 2 then
                                                response.write "Varios"
                                            else
                                                select case t("Categoria")
                                                    case "DM": response.write "Musica"
                                                    case "PE": response.write "Pelicula"
                                                    case "JU": response.write "Video Juego"
                                                    case "SO": response.write "Software"
                                                    case "LI": response.write "Libro"
                                                    case "VM": response.write "Video Musical"
                                                end select
                                            end if
                                            %></td>
                                        <td class="lista celda" style="width: 12%; text-align: left;"><%= t("Forma") %></td>
                                        <td class="lista celda" style="width: 12%; text-align: left;"><%
                                            if t("NombrePlataforma") = "- No Aplica -" then
                                                response.write "&nbsp;"
                                            else
                                                response.write t("NombrePlataforma")
                                            end if
                                        %></td>
                                        <td class="lista celda" style="width: 5%; text-align: center;"><%
                                            if t("Es_3D") = "3D" then
                                                response.write "3D"
                                            else
                                                response.write "&nbsp;"
                                            end if
                                        %></td>
                                    </tr>                                        
                                <%

                                    t.MoveNext
                                Loop until (t.eof)

                                '
                                ' Se acabaron los datos....
                                ' Presentamos los valores de la ultima ruptura
                                ' y luego el total global
                                '
                                %>
                                    <tr style = "background-color: rgb(135, 135, 135); font-family: Arial; color: rgb(255, 255, 255);">
                                        <td class="lista"  colspan="8" style = "width: 100%; padding: 10px; text-align: right; font-size: 14px;">
                                            <% 
                                                response.write "Total de Objetos:&nbsp;&nbsp;" & mediosRuptura & "&nbsp;&nbsp;<br />"
                                                response.write "Saldo:&nbsp;&nbsp;" & FormatNumber(saldoRuptura) & "&nbsp;&nbsp;"

                                                saldo = saldo + saldoRuptura
                                                medios = medios + mediosRuptura
                                            %>
                                        </td>                                               
                                    </tr>
                                <%
                            %>
                        </table>

                        <table style="width:95%;" class="center" >
                            <tr>
                                <td class="lista" style="color: white; background-color: black; font-size: 14px; text-align: right;">
                                    Total de Objetos:&nbsp;&nbsp;<%= Medios %>&nbsp;<br/>
                                    Saldo:&nbsp;&nbsp;<%= FormatNumber(Saldo) %>&nbsp;<
                                </td>
                            </tr>
                        </table>  
                    <%
                end if

                t.close: set t = nothing                
            end sub

            sub PrepararDatos()
                usu = Request.Cookies("Usuario")
                editor = Request.Form("cboEditor")
                Agrupar = Request.Form("chk_ruptura")

                for k = 0 to Request.Form("frm_Formas")
                    campo = Request.Form("f" & k)

                    if campo <> "" then
                        if listaFormas = "" then
                            listaFormas = "'" & campo & "'"
                        else
                            listaFormas = listaFormas & ", '" & campo & "'"
                        end if
                    end if
                next

                for k = 0 to Request.Form("frm_Plataformas")
                    campo = Request.Form("p" & k)

                    if campo <> "" then
                        if listaPlataformas = "" then
                            listaPlataformas = "'" & campo & "'"
                        else
                            listaPlataformas = listaPlataformas & ", '" & campo & "'"
                        end if
                    end if
                next

                select case Agrupar
                    CASE 0
                        '
                        ' No Agrupar
                        '
                        sqlString = "SELECT AEdicion, Titulo, Interprete, Precio, Casa, Categoria, Forma, NombrePlataforma, Es_3D " & _
                                    "FROM discos_Rep_Medios " & _
                                    "WHERE (Usuario = '" & usu & "') " & _
                                    "AND (CHARINDEX('" & editor & "', Categoria) > 0) " & _
                                    "AND (CodigoForma IN (" & listaFormas & ")) " 

                        if editor = "SO" OR editor = "JU" then
                            sqlString = sqlString & "AND (CodigoPlataforma IN (" & listaPlataformas & ")) " 
                        end if

                        sqlString = sqlString & "ORDER BY AEdicion, Paquete, Objeto;"
                    CASE 1
                        '
                        ' Agrupar por Plataformas
                        '
                        sqlString = "SELECT AEdicion, Titulo,  Interprete, Precio, Casa, Categoria, Forma, NombrePlataforma, Es_3D " & _
                                    "FROM  discos_Rep_Medios " & _
                                    "WHERE (Usuario = '" & usu & "') " & _
                                    "AND (CHARINDEX('" & editor & "', Categoria) > 0) " & _
                                    "AND (CodigoForma IN (" & listaFormas & ")) " 

                        if editor = "SO" OR editor = "JU" then
                            sqlString = sqlString & "AND (CodigoPlataforma IN (" & listaPlataformas & ")) " & _
                                            "ORDER BY NombrePlataforma, AEdicion, Paquete, Objeto;"
                        else
                            sqlString = sqlString & "ORDER BY AEdicion, , Paquete, Objeto;"
                        end if
                            
                    CASE 2
                        '
                        ' Agrupar por Metadata
                        '
                        sqlString = "SELECT i.AEdicion, i.Titulo, i.Interprete, i.Precio, i.Casa, i.Categoria, i.Forma, i.NombrePlataforma, i.Es_3D, m.MetaData " & _
                                    "FROM discos_Rep_Medios AS i " & _
                                "INNER JOIN discos_Paquetes_Metadata AS m " & _
                                        "ON i.Usuario = m.Usuario " & _
                                    "AND i.Paquete = m.Paquete " & _
                                    "WHERE (CHARINDEX('" & editor & "', i.Categoria) > 0) " & _
                                    "AND (i.CodigoPlataforma IN (" & listaFormas & ")) "

                        if editor = "SO" OR editor = "JU" then
                            sqlString = sqlString & "AND (CodigoPlataforma IN (" & listaPlataformas & ")) " 
                        end if

                        sqlString = sqlString & "ORDER BY m.MetaData, i.AEdicion, , Paquete, Objeto;"
                end select            
            end sub
        %>
    </head>

    <body>
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->      
        <%
            PrepararDatos      
        %>

        <br />

        <table style="width:95%;" class="center">
            <tr>
                <td class="lista"  style="width:100%; font-size: 18px; text-align:center;">
                    INFORME DE MEDIOS
                </td>
            </tr>                                   
        </table>

        <table style="width:95%;" class="titulo center">
            <tr>
                <td class="lista"  style="width:  5%; text-align: center;">Año</td>
                <td class="lista"  style="width: 30%; text-align: center;">Titulo</td>
                <td class="lista"  style="width:  8%; text-align: center;">Precio</td>
                <td class="lista"  style="width: 16%; text-align: center;">Casa</td>
                <td class="lista"  style="width: 12%; text-align: center;">Tipo</td>
                <td class="lista"  style="width: 12%; text-align: center;">Forma</td>
                <td class="lista"  style="width: 12%; text-align: center;">Plataforma</td>
                <td class="lista"  style="width:  5%; text-align: center;">3D</td>
            </tr>
        </table>

        <%
            Select Case Agrupar
                Case 0: Informe_Normal
                Case 1: Ruptura_Plataforma
                Case 2: Ruptura_Metadata
            End Select        
        %>

        <br />     

        <%
            con.close: set con = nothing
        %>   
    </body>
</html>