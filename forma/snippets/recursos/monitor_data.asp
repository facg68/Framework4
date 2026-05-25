    <%
        Dim min, run, tot, vis
        set monitorConn = Server.CreateObject("ADODB.Connection")
        monitorConn.open Application("Conn")

        monitor_SQLString = "SELECT codUsuario, codSistema, codProceso, snippet, snippetLeft, snippetTop, snippetIndex, snippetActivo, snippetMinimizado, TimeStamp " & _
                              "FROM seg_Usuarios_Snippets " & _
                             "WHERE codUsuario = '" & Request.Cookies("Usuario") & "' " & _
                          "ORDER BY snippetIndex"

        Set monitorTable = monitorConn.Execute(monitor_SQLString)
    %>

        <table class="tabla tabla-red">
            <thead>
                <th class="monitor_td sticky" style="width: 10%; text-align: center;">Usuario</th>
                <th class="monitor_td sticky" style="width: 10%; text-align: center;">Sistema</th>
                <th class="monitor_td sticky" style="width: 12%; text-align: center;">Proceso</th>
                <th class="monitor_td sticky" style="width: 20%; text-align: center;">Ventana</th>
                <th class="monitor_td sticky" style="width: 10%; text-align: center;">Pos</th>
                <th class="monitor_td sticky" style="width: 10%; text-align: center;">Capa</th>
                <th class="monitor_td sticky" style="width:  5%; text-align: center;">Act</th>
                <th class="monitor_td sticky" style="width:  5%; text-align: center;">Est</th>
                <th class="monitor_td sticky" style="width: 18%; text-align: center;">Acciones</th>                
            </thead>

            <%
                tot = 0
                min = 0
                vis = 0

                Do
                    tot = tot + 1
                    btn1_estado = ""
                    btn2_estado = ""
                    btn3_estado = ""

                    if monitorTable("snippetActivo") = 1 then vis = vis + 1
                    if monitorTable("snippetMinimizado") = 1 then min = min + 1

                %>
                    <tbody>
                        <tr>
                            <td class="monitor_td"><%= monitorTable("codUsuario") %></td>
                            <td class="monitor_td"><%= monitorTable("codSistema") %></td>
                            <td class="monitor_td"><%= monitorTable("codProceso") %></td>
                            <td class="monitor_td"><%= monitorTable("snippet") %></td>

                            <td class="monitor_td" style="text-align: center;">
                                <%= monitorTable("snippetLeft") %> ,
                                <%= monitorTable("snippetTop") %>
                            </td>

                            <td class="monitor_td" style="text-align: center;"><%= monitorTable("snippetIndex") %></td>

                            <td class="monitor_td" style="text-align: center;">
                                <% if monitorTable("snippetActivo") = 1 then %>
                                    🟢
                                <% else %>
                                    ⚫
                                <% end if %>
                            </td>

                            <td class="monitor_td" style="text-align: center;">
                                <% 
                                    if monitorTable("snippetMinimizado") = 1 then
                                        response.write "<img src='/forma/snippets/recursos/imagenes/m.png' style='height: 25px;'>"
                                    else 
                                        if monitorTable("snippetActivo") = 1 then
                                            response.write "<img src='/forma/snippets/recursos/imagenes/w.png' style='height: 25px;'>"
                                        else 
                                            response.write "&nbsp;"
                                        end if 
                                    end if 
                                %>
                            </td>

                            <td class="monitor_td" style="text-align: center;">
                                <%
                                    if monitorTable("snippetMinimizado") = 1 then btn2_estado = "disabled"

                                    if monitorTable("snippetActivo") = 0 then 
                                        btn1_estado = "disabled"
                                        btn2_estado = "disabled"
                                        btn3_estado = "disabled"
                                    end if     
                                %>

                                <button type="button" class="snippet-btn tiny verde <%= btn1_estado %>"
                                        onclick="monitor_ver('<%= monitorTable("snippet") %>')" 
                                        <%= btn1_estado %>>
                                    ver
                                </button>

                                <button type="button" class="snippet-btn tiny azul <%= btn2_estado %>" 
                                        onclick="monitor_min('<%= monitorTable("snippet") %>')" 
                                        <%= btn2_estado %>>
                                    min
                                </button>                                

                                <button type="button" class="snippet-btn small rojo <%= btn3_estado %>" 
                                        onclick="monitor_cerrar('<%= monitorTable("snippet") %>')" 
                                        <%= btn3_estado %>>
                                    cerrar
                                </button>                                
                            </td>
                        </tr>
                    </tbody>
                <%
                    monitorTable.MoveNext
                Loop Until (monitorTable.eof)
            %>

            <tfoot>
                <tr>
                    <td class="monitor_td sticky" colspan="9" style="text-align: center;">
                        Procesos: <%= tot %>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        Corriendo: <%= vis %>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        Minimizados: <%= min %>
                    </td>
                </tr>
            </tfoot>            
        </table>   
        
        <table style="padding: 0px;">
            <tr>
                <td style="text-align:left;">
                    <button type="button" class="snippet-btn violeta"
                            style="width: 100%; background-color: rgb(69, 69, 69);"
                            onclick="monitor_verTodo()">
                        Mostrar Todo
                    </button>                
                </td>

                <td style="text-align: center;">
                    <button type="button" class="snippet-btn violeta"
                            style="width: 100%; background-color: rgb(69, 69, 69);""
                            onclick="monitor_minTodo()">
                        Minimizar Todo
                    </button>                
                </td>

                <td style="text-align: right;">
                    <button type="button" class="snippet-btn violeta"
                            style="width: 100%; background-color: rgb(69, 69, 69);"
                            onclick="monitor_cerrarTodo()">
                        Cerrar Todo
                    </button>                
                </td>
            </tr>
        </table>

    <%
        monitorTable.Close: set monitorTable = nothing
        monitorConn.close: set monitorConn = nothing
    %>