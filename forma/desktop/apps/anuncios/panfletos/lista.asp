<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <title>Panfletos</title>
        <meta charset="UTF-8">
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->

        <%
            thisSystem = "anuncios"
            thisProcess = "anu.0200"
            SysLockOut

            dim cc, t, tt, sqlString, data, labels
            dim cActivas, cInactivas, estatusPanfleto, ordenadoPor

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")     
        %>          
    </head>

    <body plantilla="lista" reserva="200">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <br />

        <%
            estatusPanfleto = request.querystring("e")
            ordenadoPor = request.querystring("op")            

            if estatusPanfleto = "" then estatusPanfleto = "a"
            if ordenadoPor = "" then ordenadoPor = "0"

            cActivas = 0 
            cInactivas = 0

            sqlString = "SELECT Secuencia, Publicador, Registro, Nombre, Objeto, Desde, Hasta, " & _
                              " CreadoPor, FechaRegistro, PublicarDesde, PublicarHasta, Estado, CU " & _
                          "FROM seg_listaPanfletos " & _
                         "WHERE (Secuencia > 0) "

            select case estatusPanfleto
                case "a": sqlString = sqlString & "AND (Estado = 1) "
                case "d": sqlString = sqlString & "AND (Estado = 0) "
            end select

            select case ordenadoPor
                case 0: sqlString = sqlString & " ORDER BY Secuencia;"
                case 1: sqlString = sqlString & " ORDER BY Publicador, Secuencia;"
                case 2: sqlString = sqlString & " ORDER BY FechaRegistro;"
                case 3: sqlString = sqlString & " ORDER BY PublicarDesde;"
                case 4: sqlString = sqlString & " ORDER BY PublicarHasta;"

                case 5: sqlString = sqlString & " ORDER BY Secuencia Desc;"
                case 6: sqlString = sqlString & " ORDER BY Publicador Desc, Secuencia;"
                case 7: sqlString = sqlString & " ORDER BY FechaRegistro Desc;"
                case 8: sqlString = sqlString & " ORDER BY PublicarDesde Desc;"
                case 9: sqlString = sqlString & " ORDER BY PublicarHasta Desc;"
            end select

            set t = cc.execute(sqlString)        
        %>

        <div style="width: 98%; margin: auto;">
            <br/>

            <table style="width: 100%; margin: auto;">
                <tr style="padding: 10px;">
                    <td style="text-align:left; width: 55%;">
                        <span style="font-size: 24px">
                            <%
                                select case estatusPanfleto
                                    case "a": response.write "Panfletos Activos"
                                    case "d": response.write "Panfletos Desctivados"
                                    case else: response.write "Panfletos"
                                end select   
                            %>
                        </span>

                        <br />

                        <span style="font-size: 20px">
                            <%
                                select case ordenadoPor
                                    case 0: response.write "&nbsp;En Orden de Creacion"
                                    case 1: response.write "&nbsp;Ordenado por Publicador"
                                    case 2: response.write "&nbsp;Ordenado por Fecha de Registro"
                                    case 3: response.write "&nbsp;Ordenado por Fecha de Inicio"
                                    case 4: response.write "&nbsp;Ordenado por Fecha de Finalización"

                                    case 5: response.write "&nbsp;En Orden de Creacion (desc)"
                                    case 6: response.write "&nbsp;Ordenado por Publicador (desc)"
                                    case 7: response.write "&nbsp;Ordenado por Fecha de Registro (desc)"
                                    case 8: response.write "&nbsp;Ordenado por Fecha de Inicio (desc)"
                                    case 9: response.write "&nbsp;Ordenado por Fecha de Finalización (desc)"                                    
                                end select                            
                            %>
                        </span>
                    </td>

                    <td style="text-align:right; width: 45%;">
                        <select class="field" 
                                name="ordenadoPor" id="ordenadoPor" onChange="filtrar()">
                            <option value="0" <% if ordenadoPor = "0" then response.write " selected" %>>&#9650; Orden de Creación</option>
                            <option value="1" <% if ordenadoPor = "1" then response.write " selected" %>>&#9650; Publicador</option>
                            <option value="2" <% if ordenadoPor = "2" then response.write " selected" %>>&#9650; Fecha de Registro</option>
                            <option value="3" <% if ordenadoPor = "3" then response.write " selected" %>>&#9650; Fecha de Inicio</option>
                            <option value="4" <% if ordenadoPor = "4" then response.write " selected" %>>&#9650; Fecha de Finalización</option>

                            <option value="5" <% if ordenadoPor = "0" then response.write " selected" %>>&#9660; Orden de Creación</option>
                            <option value="6" <% if ordenadoPor = "1" then response.write " selected" %>>&#9660; Publicador</option>
                            <option value="7" <% if ordenadoPor = "2" then response.write " selected" %>>&#9660; Fecha de Registro</option>
                            <option value="8" <% if ordenadoPor = "3" then response.write " selected" %>>&#9660; Fecha de Inicio</option>
                            <option value="9" <% if ordenadoPor = "4" then response.write " selected" %>>&#9660; Fecha de Finalización</option>
                        </select>                                                 

                        <select class="field"  name="verlista" id="verlista" onChange="filtrar()">
                            <option value="*" <% if estatusPanfleto = "*" then response.write " selected" %>>Ver Todo</option>
                            <option value="a" <% if estatusPanfleto = "a" then response.write " selected" %>>Activas</option>
                            <option value="d" <% if estatusPanfleto = "d" then response.write " selected" %>>Desactivadas</option>
                        </select>                        

                        <button type="button" class="form-btn verde small" onclick="editar('*')">
                            <i class="fa fa-edit fa-xl" title="Nuevo"></i>
                        </button>                        
                    </td>                    
                </tr>
            </table>

            <table style="width: 100%; margin: auto;">
                <tr style="font-size: 14px; background-color: rgb(61, 61, 61); color:rgb(255,255,255);">
                    <td style="padding: 10px; text-align:center; width: 10%;">Estado</td>
                    <td style="padding: 10px; text-align:left;   width: 30%;">Titulo</td>
                    <td style="padding: 10px; text-align:center; width: 10%;">Desde</td>
                    <td style="padding: 10px; text-align:center; width: 10%;">Hasta</td>
                    <td style="padding: 10px; text-align:left;   width: 25%;">Publicador</td>
                    <td style="padding: 10px; text-align:center; width: 15%;">&nbsp;</td>
                </tr>

                <tr>
                    <td colspan="6">
                        <div id="overFlow" style="width:100%; height: 650px; overflow: auto; background-color: rgb(207, 207, 207);">                        
                            <table style="width: 100%;">
                                <%  
                                    if not (t.bof or t.eof) then  
                                        Do     
                                            bgcolor = "199, 230, 188"
                                            if t("Estado") = "0" then bgcolor = "245, 211, 208"
                                                CUR = t("CU") 
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

                                            <td style="padding: 5px; text-align:left;   width: 30%;" onclick="editar(<%= t("secuencia") %>)"><%= t("Nombre") %></td>
                                            <td style="padding: 5px; text-align:center; width: 10%;" onclick="editar(<%= t("secuencia") %>)"><%= t("Desde") %></td>
                                            <td style="padding: 5px; text-align:center; width: 10%;" onclick="editar(<%= t("secuencia") %>)"><%= t("Hasta") %></td>
                                            <td style="padding: 5px; text-align:left;   width: 25%;" onclick="editar(<%= t("secuencia") %>)"><%= t("Publicador") %></td>

                                            <td style="padding: 5px; text-align:right; width:15%;">
                                                <button type="button" class="form-btn verde" onclick="ver(<%= t("secuencia") %>)">
                                                    <i class=" fa fa-eye fa-xl" title="Ver panfleto"></i>
                                                </button>

                                                <button type="button" class="form-btn violeta" onclick="subirObjeto('<%= t("CU")  %>')">
                                                    <i class=" fa fa-cloud fa-xl" title="Subir Objeto"></i>
                                                </button>

                                                <button type="button" class="form-btn rojo" onclick="borrar(<%= t("secuencia") %>)">
                                                    <i class=" fa fa-trash fa-xl" title="Borrar panfleto"></i>
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
                    <td colspan="6" style="padding: 10px; text-align:center; width: 100%;">
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
                var estatus = document.getElementById("verlista").value;
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="lista.asp?e=" + estatus + "&op=" + ordenamiento;
                window.location.href = vinculo;                      
            }
                  
            function editar(panfleto) {
                var estatus = document.getElementById("verlista").value;
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="editar.asp?p=" + panfleto + "&e=" + estatus + "&op=" + ordenamiento;
                window.location.href = vinculo;
            }    

            function borrar(panfleto) {
                var confirmacion = confirm("Desea borrar el panfleto seleccionado?");
                var estatus = document.getElementById("verlista").value;
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="borrar.asp?p=" + panfleto + "&e=" + estatus + "&op=" + ordenamiento;                

                if (confirmacion) {     
                    window.location.href = vinculo;
                };
            }    

            function subirObjeto(CUR) {
                var estatus = document.getElementById("verlista").value;
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="subir_objeto.asp?cu=" + CUR + "&e=" + estatus + "&op=" + ordenamiento;    
                window.location.href = vinculo;          
            }            
                    
            function ver(panfleto) {
                var estatus = document.getElementById("verlista").value;
                var ordenamiento = document.getElementById("ordenadoPor").value;

                var vinculo ="ver.asp?p=" + panfleto + "&e=" + estatus + "&op=" + ordenamiento;                                         
                window.location.href = vinculo;          
            }   
        </script> 

        <% cc.close: set cc = nothing %>             
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>