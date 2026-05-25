<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Editar Version</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0020"
            SysLockOut
        %>
        
        <style>
            .center {
                border: none;
                margin: auto;
                width: 98%;
                padding: 0px;
            }

            .tab {
                padding: 10px; 
                border: 1px solid rgb(187, 188, 189); 
                text-align: center; 
                background-color: rgb(200, 202, 204);                
            }

			.tabDetalles {
				font-family: "Arial Narrow", Arial, sans-serif;
				font-size: 13px; 
				vertical-align: top; 
				padding: 5px; 
				border: 1px solid rgb(187, 188, 189); 
				text-align: left; 
				background-color: rgb(255, 255, 255);
				line-height:2em;
			}   
            
            .borde {
                border: 1px solid;
                border-color: rgb(184, 184, 184);
            }  
            
            .vbControl_B_Enabled {
                background-color: rgb(240, 255, 227);
                color: rgb(0, 0, 0);
                padding: 5px;
                border: 1px solid rgb(28, 69, 117);
            }   
            
            p {
                margin-bottom: 10px;
                line-height: 1.5;
            }            
        </style>

        <%
            dim con, t, p, tt, sqlString, Sistema, Version, ordenadoPor
            dim nombre, Descripcion, ClaseApp, IndiceOrdenamiento, Icono, sBitacora

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            '
            ' Funciones y Procedimientos
            '

            Function sqlQuery(Sistema, Parametro) 
                dim sqlString

                sqlString = "SELECT ConsultaSQL AS Consulta FROM seg_Parametros WHERE (Sistema = '" & Sistema & "' AND Parametro = '" & Parametro & "');"

                set ptt = con.execute(sqlString)   
                    sqlQuery = ptt("Consulta")   
                    sqlQuery = Replace(sqlQuery, "|", "'")                
                ptt.close: set ptt = nothing                
            End Function    

            Function ValorLista(Sistema, Parametro, Valor)
                dim l, sqlCommand

                sqlCommand = "SELECT Descripcion " & _
                               "FROM dbo.seg_Parametros_Valores " & _
                              "WHERE Sistema = '" & Sistema & "' " & _
                                "AND Parametro = '" & Parametro & "' " & _
                          "AND Valor = '" & Valor & "';"

                set l = con.execute(sqlCommand)
                    if not (l.bof or l.eof) then
                        ValorLista = l("Descripcion")
                    else
                        ValorLista = "- - Indefinida - -"
                    end if
                l.close: set l = nothing
            End Function

            Function ValorConsulta(Sistema, Parametro, Valor)
                dim l, sqlCommand

                ValorConsulta = " - - No Seleccionado - -"
                sqlCommand = sqlQuery(Sistema, Parametro)

                set l = con.execute(sqlCommand)
                    if not (l.bof or l.eof) then
                        do
                            if l("Codigo") = Valor then
                                ValorConsulta = l("Valor")
                            end if

                            l.MoveNext
                        Loop until (l.eof)
                    end if
                l.close: set l = nothing
            End Function            

            Sub PermisosUsuario(usuario)
                %> 
                    <div class="center">
                        <table style="width: 100%; font-size: 14px; font-family: Verdana;">
                            <tr style="background-color: rgb(85,85,85); color: rgb(255,255,255);">
                                <td style="text-align: center; padding: 10px; width: 20%;">Tipo</td>
                                <td style="text-align: center; padding: 10px; width: 55%; text-align:  left;">Nombre</td>
                                <td style="text-align: center; padding: 10px; width: 25%;">Estado</td>
                            </tr>

                            <%
                                sqlString = "SELECT CodigoUsuario, Tipo, Nombre, Estado " & _
                                            "FROM ( SELECT ru.CodigoUsuario, CASE WHEN r.TipoRol = 1 THEN 'Rol' ELSE 'Anti-Rol' END AS Tipo, " & _
                                                        " r.rolNombre AS Nombre, CASE WHEN ru.Activo = 1 THEN 'Activo' ELSE 'Desactivado' END AS Estado " & _
                                                    "FROM dbo.seg_RolesUsuarios AS ru INNER JOIN dbo.seg_Roles AS r ON ru.CodigoRol = r.rolCodigo " & _
                                            "UNION SELECT pu.CodigoUsu, CASE WHEN pu.TipoProceso = 1 THEN 'Permiso' ELSE 'Anti-Permiso' END AS Tipo, " & _
                                                        " s.sysNombre + '  |  ' + p.proNombre AS Nombre, CASE WHEN pu.Activo = 1 THEN 'Activo' ELSE 'Desactivado' END AS Estado " & _
                                                " FROM dbo.seg_ProcesosUsuarios AS pu INNER JOIN dbo.seg_Procesos AS p ON pu.CodigoProc = p.proCodigo AND pu.CodigoSis = p.proSistema " & _
                                            " INNER JOIN dbo.seg_Sistemas AS s ON p.proSistema = s.sysCodigo ) AS q " & _
                                            " WHERE CodigoUsuario = '" & Usuario & "' ORDER BY Nombre;"

                                set tt = con.execute(sqlString)

                                if not (tt.bof or tt.eof) then
                                    Do
                                        response.write "<tr>"
                                            response.write "<td style='padding: 10px; text-align: center;'>" & tt("Tipo") & "</td>"
                                            response.write "<td style='padding: 10px; text-align:  left;'>" & tt("Nombre") & "</td>"
                                            response.write "<td style='padding: 10px; text-align: center;'>" & tt("Estado") & "</td>"
                                        response.write "</tr>"                                        

                                        tt.MoveNext
                                    Loop Until tt.eof
                                end if

                                tt.close: set tt = nothing
                            %>
                        </table>
                    </div>
                <%
            end sub

            Sub ProcesosUsuario(usuario)
                dim cols, contador, k, sw, regs

                %>
                    <div class="center">                
                        <table style="width: 100%; font-size: 14px; font-family: Verdana;">
                            <%
                                sqlString = "SELECT pu.Usuario, " & _
                                                " s.sysNombre + '  |  ' + p.proNombre AS Permiso " & _
                                            "FROM dbo.seg_Sistemas AS s " & _
                                        "INNER JOIN dbo.seg_Procesos AS p " & _
                                                "ON s.sysCodigo = p.proSistema " & _
                                        "INNER JOIN dbo.seg_PermisosUsuarios AS pu " & _
                                                "ON p.proSistema = pu.Sistema " & _
                                            "AND p.proCodigo = pu.Proceso " & _
                                            "WHERE pu.Usuario = '" & usuario & "' " & _        
                                        "ORDER BY Permiso;"

                                set tt = con.execute(sqlString)

                                if not (tt.bof or tt.eof) then
                                    '
                                    ' contamos los registros...
                                    '
                                    contador = 0
                                    Do
                                        contador = contador + 1
                                        tt.MoveNext
                                    loop until tt.eof

                                    sw = -1
                                    cols = 1

                                    for k = 5 to 1 step -1
                                        if ((contador / k) - (contador \ k)) = 0 then
                                            if sw = -1 then
                                                sw = 0
                                                cols = k
                                                regs = (contador \ k)
                                            end if
                                        end if
                                    next 

                                    '
                                    ' Desplegamos los datos...
                                    '
                                    tt.moveFirst

                                    response.write "<tr style='font-size: 12px; font-family: Verdana;'>"
                                        response.write "<td style='padding: 0px;'>"

                                            response.write "<table style='width: 100%; font-size: 14px; font-family: Verdana;'>"
                                                response.write "<tr style='font-size: 12px; font-family: Verdana;'>"

                                                    for k = 1 to cols
                                                        response.write "<td style='padding: 10px; vertical-align: top;'>" 
                                                            for r = 1 to regs
                                                                response.write "<p>" & tt("Permiso") & "</p>"
                                                                tt.MoveNext
                                                            next 
                                                        response.write "</td>"
                                                    next 

                                                response.write "</tr>"    
                                            response.write "</table>"

                                        response.write "</td>"
                                    response.write "</tr>"
                                end if

                                tt.close: set tt = nothing
                            %>
                        </table>
                    </div>
                <%
            end Sub

            Sub VariablesUsuario(usuario)
                %> 
                    <div class="center">
                        <table style="width: 100%; font-size: 14px; font-family: Verdana;">
                            <tr style="background-color: rgb(85,85,85); color: rgb(255,255,255);">
                                <td style="text-align: center; padding: 10px; width: 20%;">Sistema</td>                            
                                <td style="text-align: center; padding: 10px; width: 20%;">Variable</td>
                                <td style="text-align: center; padding: 10px; width: 40%;">Descripcion</td>
                                <td style="text-align: center; padding: 10px; width: 20%;">Valor</td>
                            </tr>

                            <%
                                sqlString = "SELECT s.sysCodigo, s.sysNombre AS Sistema, pu.Parametro, p.TipoParametro AS Tipo, p.Descripcion, pu.Valor " & _
                                              "FROM dbo.seg_Usuarios_Parametros AS pu " & _
                                        "INNER JOIN dbo.seg_Parametros AS p " & _
                                                "ON pu.Parametro = p.Parametro " & _
                                               "AND pu.Sistema = p.Sistema " & _
                                        "INNER JOIN dbo.seg_Sistemas AS s " & _
                                                "ON p.Sistema = s.sysCodigo " & _
                                             "WHERE (pu.Usuario = '" & usuario & "') " & _
                                          "ORDER BY Sistema, Parametro;"

                                set tt = con.execute(sqlString)

                                if not (tt.bof or tt.eof) then
                                    Do
                                        response.write "<tr>"
                                            response.write "<td style='padding: 15px; width: 20%; text-align: center;'>" & tt("Sistema") & "</td>"
                                            response.write "<td style='padding: 15px; width: 20%; text-align: center;'>" & tt("Parametro") & "</td>"
                                            response.write "<td style='padding: 15px; width: 40%; ext-align: center;'>" & tt("Descripcion") & "</td>"
                                            response.write "<td style='padding: 15px; width: 20%; text-align: center;"
                                                if tt("Tipo") = 6 then
                                                    response.write " border: 1px dashed rgb(82, 82, 82);"
                                                    response.write " background-color: " & tt("Valor") & ";"
                                                end if
                                            response.write "'>"
                                                Select Case tt("Tipo")
                                                    Case "1" 'Permiso'
                                                        response.write "&nbsp;"
                                                    Case "2" 'Variable'
                                                        response.write tt("Valor")
                                                    Case "3" 'Sí / No'
                                                        if tt("Valor") = 1 then
                                                            response.write "Sí"
                                                        else
                                                            response.write "No"
                                                        end if
                                                    Case "4" 'Lista'
                                                        response.write ValorLista(tt("sysCodigo"), tt("Parametro"), tt("Valor"))
                                                    Case "5" 'Consulta SQL'
                                                        response.write ValorConsulta(tt("sysCodigo"), tt("Parametro"), tt("Valor"))
                                                    Case "6" 'Color'
                                                        response.write tt("Valor")   
                                                    Case 7
                                                        %>
                                                            <label style="width: 100%;">
                                                                <input class=" form-control" style="width: 100%" type="range" min="0" max="100" value="<%= tt("Valor") %>" disabled>
                                                            </label>                                                                                                         
                                                        <%
                                                End Select
                                            response.write "</td>"
                                        response.write "</tr>"                                        

                                        tt.MoveNext
                                    Loop Until tt.eof
                                end if

                                tt.close: set tt = nothing
                            %>
                        </table>
                    </div>
                <%            
            end Sub

            Sub BitacoraUsuario(Usuario)
                %> 
                    <div class="center">
                        <table style="width: 100%; font-size: 14px; font-family: Verdana;">
                            <tr style="background-color: rgb(85,85,85); color: rgb(255,255,255);">
                                <td style="text-align: left; padding: 10px; width: 30%;">Usuario</td>
                                <td style="text-align: left; padding: 10px; width: 30%;">Sistema</td>
                                <td style="text-align: left; padding: 10px; width: 25%;">Proceso</td>
                                <td style="text-align: center; padding: 10px; width: 15%;">Fecha</td>
                            </tr>

                            <%
                                sqlString = "SELECT Amo, Mes, Fecha, Sistema, Proceso, Usuario, Acceso, FechaForm, nomUsuario, nomSistema, nomProceso " & _
                                              "FROM dbo.qry_seg_Bitacoras AS b " & _
                                            " WHERE Usuario = '" & Usuario & "' ORDER BY Fecha DESC;"

                                set tt = con.execute(sqlString)

                                if not (tt.bof or tt.eof) then
                                    Do
                                        response.write "<tr>"
                                            response.write "<td style='padding: 10px; text-align:  left;'>" & tt("nomUsuario") & "</td>"
                                            response.write "<td style='padding: 10px; text-align:  left;'>" & tt("nomSistema") & "</td>"
                                            response.write "<td style='padding: 10px; text-align:  left;'>" & tt("nomProceso") & "</td>"
                                            response.write "<td style='padding: 10px; text-align: center;'>" & tt("FechaForm") & "</td>"
                                        response.write "</tr>"                                        

                                        tt.MoveNext
                                    Loop Until tt.eof
                                end if

                                tt.close: set tt = nothing
                            %>
                        </table>
                    </div>
                <%
            end Sub

            function FechaFormulario(FechaSQL)       
                dim d, m, a, h, mm

                if FechaSQL <> "" then
                    a = left(FechaSQL, 4)
                    m = right("00" & mid(FechaSQL, 6, 2), 2)
                    d = right("00" & mid(FechaSQL, 9, 2), 2)

                    h = right("00" & mid(FechaSQL, 12, 2), 2)
                    mm = right("00" & mid(FechaSQL, 15, 2), 2)

                    FechaFormulario = d & "/" & m & "/" & a & " " & h & ":" & mm
                else
                    FechaFormulario = ""
                end if
            end function    

            function SmallFechaFormulario(FechaSQL)       
                dim d, m, a

                if FechaSQL <> "" then
                    a = left(FechaSQL, 4)
                    m = right("00" & mid(FechaSQL, 6, 2), 2)
                    d = right("00" & mid(FechaSQL, 9, 2), 2)

                    SmallFechaFormulario = d & "/" & m & "/" & a 
                else
                    SmallFechaFormulario = ""
                end if
            end function                                 
        %>
    </head>

    <body plantilla="normal" reserva="165" onload="Tabs_Display('tab_procesos')">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->  

        <br /> 

        <%
            usuario = request.querystring("u")
            ordenadoPor = request.querystring("o")
        %>

        <div style="width: 95%; margin: auto;">
            <%
                sqlString = "SELECT * " & _
                            "FROM seg_Usuarios " & _
                            "WHERE (usuCodigo = '" & Usuario & "');"
                
                set t = con.execute(sqlString)                
            %>
      
            <table style="width: 95%;">
                <tr>
                    <td style="width: 30%; font-size: 20px;">
                        <h3><%
                            response.write "Datos de " & t("usuNombre") 
                            %>
                        </h3>
                    </td>

                    <td style="width: 70%; text-align: right;">
                        <button class='form-btn rojo normal'  type='button' onclick="volver()">Cancelar</button>&nbsp;&nbsp;                         
                    </td>
                </tr>
            </table>    

            <div class="main">
                <div class="line">
                    <label class="label full section">
                        <table style="width: 100%; border: none; border-spacing: 0px;">
                            <tr>
                                <td style="padding: 10px; border: 1px solid rgb(187, 188, 189); text-align: center; background-color: rgb(200, 202, 204);"
                                    onclick="Tabs_Display('tab_procesos')">
                                    Procesos Asignados
                                </td>

                                <td style="padding: 10px; border: 1px solid rgb(187, 188, 189); text-align: center; background-color: rgb(200, 202, 204);"
                                    onclick="Tabs_Display('tab_permisos')">
                                    Roles y Permisos
                                </td>

                                <td style="padding: 10px; border: 1px solid rgb(187, 188, 189); text-align: center; background-color: rgb(200, 202, 204);"
                                    onclick="Tabs_Display('tab_variables')">
                                    Variables Asignadas
                                </td>                                                    

                                <td style="padding: 10px; border: 1px solid rgb(187, 188, 189); text-align: center; background-color: rgb(200, 202, 204);"
                                    onclick="Tabs_Display('tab_bitacora')">
                                    Bitácora de Accesos
                                </td>                                                        
                            </tr>

                            <tr>
                                <td colspan="4" style="padding: 10px; border: 1px solid rgb(187, 188, 189); text-align: center; background-color: rgb(255, 255, 255);">
                                    <div id="tab_procesos" style="display: none; text-align: left; font-size: 16px; line-height: 1.8em;">
                                        <% ProcesosUsuario usuario %>
                                    </div>

                                    <div id="tab_permisos" style="display: none; text-align: left; font-size: 16px; line-height: 1.8em;">
                                        <% PermisosUsuario usuario %>
                                    </div>

                                    <div id="tab_variables" style="display: none; text-align: left; font-size: 16px; line-height: 1.8em;">
                                        <% VariablesUsuario usuario %>
                                    </div>                                                        

                                    <div id="tab_bitacora" style="display: none; text-align: left; font-size: 16px; line-height: 1.8em;">
                                        <% BitacoraUsuario Usuario %>
                                    </div>      
                                </td>
                            </tr>
                        </table>
                    </label>
                </div>
            </div>
        </div>

        <br /><br />

        <script>
            function volver() {
                var vinculo = "lista.asp?o=<%= ordenadoPor %>";
                window.location.href = vinculo;
            }

            function Tabs_Display(codigo) {
                var t1 = document.getElementById("tab_procesos");
                var t2 = document.getElementById("tab_permisos");
                var t3 = document.getElementById("tab_variables");
                var t4 = document.getElementById("tab_bitacora");

                t1.style.display = 'none';
                t2.style.display = 'none';
                t3.style.display = 'none';
                t4.style.display = 'none';

                switch (codigo) {
                    case "tab_procesos":
                        t1.style.display = 'block';
                        break;
                    case "tab_permisos":
                        t2.style.display = 'block';               
                        break;
                    case "tab_variables":
                        t3.style.display = 'block';               
                        break;                        
                    case "tab_bitacora":
                        t4.style.display = 'block';               
                        break;                                                
                }
            }
        </script> 

        <%
            t.close: set t = nothing
            con.close: set con = nothing
        %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->        
    </body>
</html>