<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Asignar Pantallas</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "anuncios"
            thisProcess = "anu.0120"
            SysLockOut
        %>
        
        <style>
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
        </style>

        <%
            dim pCon, ptt, pProc, sqlString, Sistema, Proceso, ordenadoPor, tt
            dim Nombre, Descripcion, TipoRol, cuantosSistemas, primerTab

            Pantalla = request.querystring("p")

            set pCon = Server.CreateObject("ADODB.Connection")
            pCon.open Application("Conn")

            function NombrePantalla(Pantalla)
                dim sqlString

                sqlString = "SELECT Nombre FROM seg_Anuncios_Pantallas WHERE (Pantalla ='" & Pantalla & "');"

                set ptt = pCon.execute(sqlString)   
                    NombrePantalla = ptt("Nombre") 
                ptt.close: set ptt = nothing
            end function   

            function CuantosUsuariosQuedan(Pantalla)
                sqlString = "SELECT COUNT(*) AS CUANTOS FROM (" & _
                            "SELECT DISTINCT TOP (100) PERCENT u.usuCodigo, u.usuNombre " & _
                            "FROM dbo.seg_PermisosUsuarios AS pu INNER JOIN dbo.seg_Usuarios AS u ON pu.Usuario = u.usuCodigo " & _
                            "WHERE (pu.Sistema = 'anuncios') AND (pu.Proceso = 'anu.0100' OR pu.Proceso = 'anu.0130') " & _
                            "AND (pu.Usuario NOT IN (SELECT Usuario FROM dbo.seg_Anuncios_Pantallas_Usuarios WHERE (Pantalla = '" & Pantalla & "'))) " & _
                            "ORDER BY u.usuNombre) AS t;"

                set ptt = pCon.execute(sqlString)   
                    CuantosUsuariosQuedan = ptt("Cuantos") 
                ptt.close: set ptt = nothing
            end function

            sub UsuariosAsignados(Sistema, Proceso)
                %>
                    <table style="width: 100%;" style="padding: 0px;"> 
                        <tr style="background-color: rgb(0, 0, 0); color: rgb(255, 255, 255);">
                            <td class="top borde" style="width:  5%; text-align: center; padding: 5px;">&nbsp;</td>
                            <td class="top borde" style="width: 45%; text-align: left;   padding: 5px;">Usuario</td>
                            <td class="top borde" style="width: 45%; text-align: left;   padding: 5px;">Cargo</td>
                            <td class="top borde" style="width:  5%; text-align: center; padding: 5px;">&nbsp;</td>              
                        </tr>

                        <% if CuantosUsuariosQuedan(Pantalla) > 0 then %>
                            <tr>
                                <td class="borde" style="width: 5%; text-align: center; font-size: 30px; font-weight: bold;">+</td>              

                                <td colspan="2" class="borde" style="width: 55%;">
                                    <%
                                        sqlString = "SELECT DISTINCT TOP (100) PERCENT u.usuCodigo, u.usuNombre " & _
                                                    "FROM dbo.seg_PermisosUsuarios AS pu INNER JOIN dbo.seg_Usuarios AS u ON pu.Usuario = u.usuCodigo " & _
                                                    "WHERE (pu.Sistema = 'anuncios') AND (pu.Proceso = 'anu.0100' OR pu.Proceso = 'anu.0130') " & _
                                                    "AND (pu.Usuario NOT IN (SELECT Usuario FROM dbo.seg_Anuncios_Pantallas_Usuarios WHERE (Pantalla = '" & Pantalla & "'))) " & _
                                                    "ORDER BY u.usuNombre;"                                                    

                                        set tt = pCon.execute(sqlString)
                                            response.write "<select name='usuNuevo' id='usuNuevo' class='form-control vbControl_B_Enabled'>"
                                                response.write "<option value='*'> - - Seleccione un Usuario - - </option>"

                                                Do
                                                    response.write "<option value='" & tt("usuCodigo") & "'>" & tt("usuNombre") & "</option>"
                                                    tt.MoveNext
                                                Loop Until tt.eof

                                            response.write "/<select>"
                                        tt.close: set tt = nothing
                                    %>                                    
                                </td>              

                                <td class="borde" style="width: 5%; text-align: center; background-color: rgb(240, 255, 227);">
                                    <button type="button" class="form-btn verde" onclick="NuevaLinea()"><i class="fa fa-save fa-xl"></i></button>
                                </td>              
                            </tr>    
                        <% end if %>

                        <%
                            sqlString = "SELECT p.Pantalla, p.Usuario, u.usuNombre, u.usuCargo " & _
                                          "FROM dbo.seg_Anuncios_Pantallas_Usuarios AS p " & _
                                    "INNER JOIN dbo.seg_Usuarios AS u " & _
                                            "ON p.Usuario = u.usuCodigo " & _
                                         "WHERE (p.Pantalla ='" & Pantalla & "') " & _
                                      "ORDER BY u.usuNombre;"

                            set cbox = pCon.execute(sqlString)
                                if not (cbox.bof or cbox.eof) then
                                    contador = 0

                                    Do        
                                        contador = contador + 1

                                        response.write "<tr style='background-color: rgb(255,255,255);'>"

                                            response.write "<td style='text-align: center; padding: 5px; width: 5%' class='borde'>"
                                                response.write contador
                                            response.write "</td>"      

                                            response.write "<td style='text-align: left; padding: 5px; width: 25%' class='borde'>"
                                                response.write cbox("usuNombre")
                                            response.write "</td>"                                   

                                            response.write "<td style='text-align: left; padding: 5px; width: 25%' class='borde'>"
                                                response.write cbox("usuCargo")
                                            response.write "</td>"

                                            response.write "<td style='text-align: center; padding: 5px; width: 5%' class='borde'>"
                                                %><button class="form-btn rojo" type="button" onClick="BorrarPantallaAsignada('<%= cbox("Usuario") %>')"><%
                                                    response.write "<i class=' fa fa-trash fa-xl' title='Borrar Asignacion'></i>"
                                                response.write "</button>"                    
                                            response.write "</td>"                
                                        response.write "</tr>"

                                        cbox.MoveNext
                                    Loop Until (cbox.eof)
                                end if
                            cbox.close: set cbox = nothing     
                        %>                   
                    </table> 
                <%     
            end sub                  
        %>
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->   

        <%
            set ptt = pCon.execute("SELECT Nombre, Descripcion FROM seg_Anuncios_Pantallas WHERE Pantalla = '" & Pantalla & "';")
                Nombre = ptt("Nombre") & " (" & ptt("Descripcion") & ")"
            ptt.close: set ptt = nothing
        %>

        <div style="width 100%; margin: auto;">
            <br />

            <table style="width: 95%; margin: auto;">
                <tr>
                    <td style="width: 70%;">
                        <%
                            response.write "<span style='font-size: 22px'>" & Nombre & "</span><br />"
                        %>
                    </td>

                    <td style="width: 30%; text-align: right;">
                        <button type='button' class='form-btn verde' style="width:100px; font-size: 16px; color: white;" onclick="volver()">Volver</button>&nbsp;&nbsp;
                    </td>
                </tr>
            </table>    

            <div class="main main-scroll">
                <form id="formulario"  name="formulario" method="post">
                    <div class="line">
                        <label class="label small">Usuarios Asignados</label>

                        <label class="label section full">   
                            <% UsuariosAsignados Sistema, Proceso %> 
                        </label>
                    </div>
                </form>
            </div>
        </div>

        <script>
            function volver() {
                var vinculo = "pantallas.asp?s=<%= Sistema %>&o=<%= ordenadoPor %>";
                window.location.href = vinculo;
            }

            function NuevaLinea() {
                var usuNuevo = document.getElementById("usuNuevo").value;

                if (usuNuevo == "*") {
                    window.alert("Debe Seleccionar un Usuario");
                } else {
                    var vinculo ="pantallas_asignar_usuario.asp?p=<%= Pantalla %>" + "&u=" + usuNuevo;
                    window.location.href = vinculo;    
                }
            }

            function BorrarPantallaAsignada(usuario) {
                var confirmacion = confirm("Desea quitar de este permiso al usuario seleccionado?");
                var vinculo ="pantallas_asignar_borrar.asp?p=<%= Pantalla %>&u=" + usuario;

                if (confirmacion) {     
                    window.location.href = vinculo;
                };
            }    
        </script> 

        <% pCon.close: set pCon = nothing %>
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>