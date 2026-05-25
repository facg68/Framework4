<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0090"
            SysLockOut
        %> 
         
        <style>
            html { background-color: black; } 

            body, html { margin: 0; padding: 0; height: 100%; overflow: auto; } 

            table, th, td {
                border-collapse: collapse;
            }

            td {
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
            
            tr:not(:last-child) { border: none !important; }
        </style>

        <%
            function FechaPagina(FechaServer)
                dim dia, mes, amo

                if FechaServer = "" then
                    FechaPagina = NULL
                else
                    amo = left(FechaServer, 4)
                    mes = right("00" & mid(FechaServer, 6, 2), 2)
                    dia = right("00" & mid(FechaServer, 9, 2), 2)

                    FechaPagina = dia & "/" & mes & "/" & amo
                end if
            end function

            function Moneda(localMonetario)
                dim cc, tt, sqlString

                sqlString = "SELECT MonedaEnteraUnica AS Nombre " & _
                            "FROM seg_cripto_NumParse_Locales " & _
                            "WHERE Local = '" & localMonetario &"';"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.execute(sqlString)

                if (tt.bof or tt.eof) then
                    Moneda = Null
                else
                    Moneda = tt("Nombre")
                end if

                tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function

            function limpiar(cadena)
                dim char, aasc, k

                limpiar = ""

                for k = 1 to len(cadena)
                    char = mid(cadena, k, 1)
                    aasc = asc(char)

                    if aasc = 32 then
                        limpiar = limpiar & "&nbsp;"
                    else
                        limpiar = limpiar & char
                    end if
                next
            end function    

            function CuantasListas(Usuario, Presupuesto)
                dim sqlString, cc, tt

                sqlString = "SELECT COUNT(*) AS Cuantos " & _
                                "FROM pre_Presupuesto_Detalles AS d " & _
                                "WHERE (Usuario = '" & Usuario & "') " & _ 
                                "AND (Presupuesto = '" & Presupuesto & "') " & _
                                "AND (LEFT(CuentaDestino, 3) = '*L:');"   

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.execute(Sqlstring)

                    CuantasListas = tt("Cuantos")

                tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function  

            sub PresentarListas(Usuario, Presupuesto)     
                '
                ' Las Listas siempre apaecerán en el panel derecho
                ' de la página
                '
                dim sqlString, t, con

                sqlString = "SELECT FORMAT(Fecha, 'dd/MM/yyyy', 'en-US') AS F, FORMAT(Hora, '##:##') AS H, " & _
                                  " RIGHT(CuentaDestino, LEN(CuentaDestino) - 3) AS Lista, Descripcion " & _
                              "FROM pre_Presupuesto_Detalles AS d " & _
                             "WHERE (Usuario = '" & Usuario & "') " & _ 
                               "AND (Presupuesto = '" & Presupuesto & "') " & _
                               "AND (LEFT(CuentaDestino, 3) = '*L:') " & _
                          "ORDER BY Fecha, Hora;"

                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")
                set t = con.execute(Sqlstring)

                if not (t.eof or t.bof) then
                    Do
                        response.write "<br /><span style='font-family: Tahoma; font-size: 14px; font-weight: bold;'>"
                        response.write t("F") & "&nbsp-&nbsp;" & t("H") & "&nbsp;|&nbsp;" & limpiar(t("Descripcion"))
                        response.write "</span><br /><br />"

                        DibujarLista usu, t("Lista")

                        response.write "<br /><br />"

                        t.MoveNext
                    Loop Until t.eof
                end if

                t.close: set t = nothing    
                con.close: set con = nothing                        
            end sub

            sub DibujarLista(Usuario, Lista)        
                '
                ' Solo pueden dibujarse las Listas que son una Cuenta
                '
                dim cc, tt, sqlString

                sqlString = "SELECT e.Codigo, e.Nombre, e.Contacto, e.PrecioOriginal, " & _
                                  " e.PrecioFinal, d.Fecha, d.Item, d.PrecioOriginal AS Monto1, d.Precio AS Monto2 " & _
                              "FROM pre_Listas_Encabezado AS e " & _
                        "INNER JOIN pre_Listas_Detalles AS d " & _
                                "ON e.Usuario = d.Usuario " & _
                               "AND e.Codigo = d.Codigo " & _
                             "WHERE (e.Usuario = '" & Usuario & "') " & _
                               "AND (e.Cuenta = 1) " & _
                               "AND (e.VerListaEnInforme = 1) " & _
                               "AND (e.Codigo = '" & Lista & "') " & _
                          "ORDER BY d.Fecha, d.Item;"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.execute(sqlString)

                if not (tt.bof or tt.eof) then 
                    sw = -1
                    total = 0
                %>
                    <table style="width: 100%;">
                        <tr class="titulo" style="text-align: left; font-weight: bold;">
                            <td colspan="2">
                                <%= limpiar(tt("Nombre")) %>
                            </td>
                        </tr>

                        <%
                            Do
                                sw = -1 * sw
                                total = total + tt("Monto1")

                                response.write "<tr class='"
                                    if sw = 1 then 
                                        response.write "impar"
                                    else
                                        response.write "par"
                                    end if
                                response.write "'>"
                                    response.write "<td class='celda' style='width: 75%;'>" & limpiar(tt("Item")) & "</td>"
                                    response.write "<td class='celda' style='width: 25%; text-align:right;'>" & FormatNumber(tt("Monto1")) & "</td>"
                                response.write "</tr>"

                                tt.MoveNext
                            Loop Until tt.eof
                        %>

                        <tr class="titulo">
                            <td colspan="2" style="text-align: right;">
                                <%= "Total:&nbsp;&nbsp;" & FormatNumber(Total) & "&nbsp;" %>
                            </td>
                        </tr>                        
                    </table>
                <%
                end if

                tt.close: set tt = nothing
                cc.close: set cc = nothing
            end sub

            function CuantasNotasDireccionadas(Usuario, Presupuesto, Direccion)
                dim sqlString, cc, tt

                sqlString = "SELECT COUNT(*) AS Cuantos FROM pre_Presupuesto_Detalles " & _
                            "WHERE (Usuario = '" & Usuario & "') AND (Presupuesto = '" & Presupuesto & "') " & _
                            "AND (Nota <> '') AND (NotaDonde = '" & Direccion & "');"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.execute(Sqlstring)

                    CuantasNotasDireccionadas = tt("Cuantos")

                tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function

            sub PresentarNotasDireccionadas(Usuario, Presupuesto, Direccion)
                dim sqlString, t, con

                sqlString = "SELECT FORMAT(Fecha, 'dd/MM/yyyy', 'en-US') AS F, FORMAT(Hora, '##:##') AS H, " & _ 
                                  " Descripcion, Nota, NotaPre " & _ 
                             " FROM pre_Presupuesto_Detalles " & _
                             "WHERE (Usuario = '" & Usuario & "') AND (Presupuesto = '" & Presupuesto & "') " & _
                               "AND (Nota <> '') AND (NotaDonde = '" & Direccion & "') " & _
                          "ORDER BY Fecha, Hora;"

                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")
                set t = con.execute(Sqlstring)

                if not (t.eof or t.bof) then
                    Do
                        response.write "<br /><span style='font-family: Tahoma; font-size: 14px; font-weight: bold;'>"
                        response.write t("F") & "&nbsp-&nbsp;" & t("H") & "&nbsp;|&nbsp;" & limpiar(t("Descripcion"))
                        response.write "</span><br /><br />"
                                            
                        if t("NotaPre") = 1 then
                            response.write "<pre>" & Replace(t("Nota"), vbCrLf, "<br />") & "</pre>"
                        else
                            '
                            ' Las Notas "normales" pueden formatearse de forma compleja, 
                            ' por lo que se asume que tendrán identificadores
                            '                  
                            response.write t("Nota") 
                        end if      

                        response.write "<br /><br />"                             

                        t.MoveNext
                    Loop Until t.eof
                end if

                t.close: set t = nothing  
                con.close: set con = nothing      
            end sub 

            function FormatCeroNull(Valor)
                if Valor <> 0 then
                    FormatCeroNull = FormatNumber(Valor)
                else
                    FormatCeroNull = "&nbsp;"
                end if
            end function
        %>
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    
        <br />  
        <%
            dim con, p, d, sqlString, pre, usu, sw, t

            usu = Request.Cookies("Usuario")
            pre = Request.QueryString("p")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            set p = con.execute("SELECT Presupuesto, Tipo, Nombre, Desde, Hasta, MultiPrecio, MonedaOrigen, MonedaDestino " & _
                                "FROM pre_Presupuesto_Encabezado " &_ 
                                "WHERE (Usuario = '" & usu & "') " & _
                                "AND (Presupuesto = '" & pre & "');")
            
            if not (p.bof or p.eof) then
        %>
        
        <!-- ENCABEZADO DE PRESUPUESTO -->

        <table style="width:95%;" class="center">
            <tr>
                <td style="width:100%; font-size: 18px; text-align:center;">
                    <%= p("Nombre") %>
                </td>
            </tr>

            <tr>
                <td style="width:100%; font-size: 14px; text-align:center;">
                    <%
                        response.write "Desde&nbsp;" & FechaPagina(p("Desde")) & "&nbsp;hasta&nbsp;" & FechaPagina(p("Hasta")) 
                    %>
                </td>
            </tr>                                        
        </table>

        <!--DETALLES DE PRESUPUESTO -->

        <%
            sqlString = "exec pre_Reporte_Detalles2 '" & usu &"', '" & pre & "';"

            set d = con.execute(sqlString)

            if not (d.eof or d.bof) then
                sw = -1
                SaldoC = 0
                SaldoE = 0
                SaldoReal = 0
                Saldo = 0
        %>
                <table style="width:95%;" class="center">
                    <%  
                        if p("MultiPrecio") = 0 then 
                                response.write "<tr class='titulo center'>"
                                    response.write "<td style='width: 12%;'>Fecha</td>"
                                    response.write "<td style='width: 40%;'>Descripcion</td>"
                                    response.write "<td style='width: 12%;'>" & Moneda(p("MonedaOrigen")) & "</td>"
                                    response.write "<td style='width: 12%;'>Saldo</td>"
                                    response.write "<td style='width: 12%;'>Efectivo</td>"
                                    response.write "<td style='width: 12%;'>S.Efec.</td>"
                                response.write "</tr>"
                        else 
                                response.write "<tr class='titulo center'>"
                                    response.write "<td style='width: 10%;'>Fecha</td>"
                                    response.write "<td style='width: 30%;'>Descripcion</td>"
                                    response.write "<td style='width: 10%;'>" & Moneda(p("MonedaOrigen")) & "</td>"
                                    response.write "<td style='width: 10%;'>" & Moneda(p("MonedaDestino")) & "</td>"
                                    response.write "<td style='width: 10%;'>Saldo</td>"
                                    response.write "<td style='width: 10%;'>Efectivo</td>"
                                    response.write "<td style='width: 10%;'>Ef.Cambio</td>"
                                    response.write "<td style='width: 10%;'>S.Efec.</td>"
                                response.write "</tr>"
                        end if

                        Do 
                            sw = -1 * sw

                            SaldoC = SaldoC + d("CarteraMonto") 
                            SaldoE = SaldoE + d("EfectivoMonto")
                            if d("Aplicado") = 1 then SaldoReal = SaldoReal + d("CarteraMonto") + d("EfectivoMonto")
                            Saldo = Saldo + d("CarteraMonto") + d("EfectivoMonto")

                            if p("MultiPrecio") = 0 then
                    %>
                                <tr class="<% if sw = 1 then response.write "par" else response.write "impar"%>">
                                    <td class="celda" style="width: 12%; text-align: center;"><%= d("FechaHora") %></td>
                                    <td class="celda" style="width: 40%;">
                                        <input type="checkbox" <%
                                            if d("Aplicado") = 1 then
                                                response.write " checked='checked' "
                                            end if
                                        %> disabled>&nbsp
                                        <%= limpiar(d("Descripcion")) %>
                                    </td>
                                    <td class="celda" style="width: 12%; text-align: right;"><%= FormatCeroNull(d("CarteraMonto")) %></td>
                                    <td class="celda" style="width: 12%; text-align: right;"><%= FormatCeroNull(SaldoC) %></td>                                    
                                    <td class="celda" style="width: 12%; text-align: right;"><%= FormatCeroNull(d("EfectivoMonto")) %></td>
                                    <td class="celda" style="width: 12%; text-align: right;"><%= FormatCeroNull(SaldoE) %></td>
                                </tr>
                    <%
                            else
                    %>
                                <tr class="<% if sw = 1 then response.write "par" else response.write "impar"%>">
                                    <td class="celda" style="width: 10%; text-align: center;"><%= d("FechaHora") %></td>
                                    <td class="celda" style="width: 30%;">
                                        <input type="checkbox" <%
                                            if d("Aplicado") = 1 then
                                                response.write " checked='checked' "
                                            end if
                                        %> disabled>&nbsp
                                        <%= limpiar(d("Descripcion")) %>
                                    </td>
                                    <td class="celda" style="width: 10%; text-align: right;"><%= FormatCeroNull(d("CarteraMonto")) %></td>
                                    <td class="celda" style="width: 10%; text-align: right;"><%= FormatCeroNull(d("CarteraCambio")) %></td>
                                    <td class="celda" style="width: 10%; text-align: right;"><%= FormatCeroNull(SaldoC) %></td>                                    
                                    <td class="celda" style="width: 10%; text-align: right;"><%= FormatCeroNull(d("EfectivoMonto")) %></td>
                                    <td class="celda" style="width: 10%; text-align: right;"><%= FormatCeroNull(d("EfectivoCambio")) %></td>
                                    <td class="celda" style="width: 10%; text-align: right;"><%= FormatCeroNull(SaldoE) %></td>
                                </tr>                                        
                    <%
                            end if

                            d.MoveNext
                        Loop until (d.eof)

                        if p("MultiPrecio") = 0 then 
                            response.write "<tr class='titulo center'>"
                                response.write "<td colspan='6' style='text-align: right;'>Balance Segun Transacciones:&nbsp;&nbsp;" & FormatNumber(Saldo) & "&nbsp;</td>"
                            response.write "</tr>"

                            response.write "<tr class='titulo center'>"
                                response.write "<td colspan='6' style='text-align: right;'>Saldo Real:&nbsp;&nbsp;" & FormatNumber(SaldoReal) & "&nbsp;</td>"
                            response.write "</tr>"                            
                        else
                            response.write "<tr class='titulo center'>"
                                response.write "<td colspan='8' style='text-align: right;'>Balance Segun Transacciones:&nbsp;&nbsp;" & FormatNumber(Saldo) & "&nbsp;</td>"
                            response.write "</tr>"                        

                            response.write "<tr class='titulo center'>"
                                response.write "<td colspan='8' style='text-align: right;'>Saldo Real:&nbsp;&nbsp;" & FormatNumber(SaldoReal) & "&nbsp;</td>"                            
                            response.write "</tr>"                        
                        end if
                    %>
                </table>
        <%
            end if
        %>


        <!-- Tablas Auxiliares y Notas -->

        <table style="width:95%;" class="center">
            <tr>

            <%
                if (CuantasNotasDireccionadas(usu, pre, "I") > 0) AND ((CuantasListas(usu, pre) + CuantasNotasDireccionadas(usu, pre, "D")) > 0) then
                    '
                    ' Normal... Presentamos ambas partes...
                    '
                    %>
                        <td style="width: 50%; vertical-align: top;">
                            <!-- Seccion de Notas -->

                            <div style="font-family: 'Courier New', monospace; font-size: 12px; padding:10px; width:95%;">
                                <% PresentarNotasDireccionadas usu, pre, "I" %>
                            </div>
                        </td>
                        
                        <td style="width: 50%; vertical-align: top;">
                            <!-- Seccion de Listas -->
                                                       
                            <% 
                                PresentarListas usu, pre 
                                PresentarNotasDireccionadas usu, pre, "D" 
                            %>                            
                        </td>                      
                    <%
                else
                    '
                    ' Ok... Veamos cual de las dos será presentada
                    '
                    if CuantasNotasDireccionadas(usu, pre, "I") > 0 then
                        '
                        ' Presentamos las Notas
                        '
                        %>
                            <td style="width: 100%; vertical-align: top;">
                                <!-- Seccion de Notas -->

                                <div style="font-family: 'Courier New', monospace; font-size: 12px; padding:10px; width:95%;">
                                    <% PresentarNotasDireccionadas usu, pre, "I" %>
                                </div>
                            </td>                        
                        <%
                    else
                        '
                        ' Presentamos las Listas
                        '
                        %>
                            <td style="width: 100%; vertical-align: top;">
                                <!-- Seccion de Listas -->
                                
                                <% 
                                    PresentarListas usu, pre 
                                    PresentarNotasDireccionadas usu, pre, "D"
                                %>
                            </td>                           
                        <%
                    end if
                end if
            %>                                         
            </tr>
        </table>

        <br />     

        <%
            end if

            p.close: set p = nothing
            con.close: set con = nothing
        %>   
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>