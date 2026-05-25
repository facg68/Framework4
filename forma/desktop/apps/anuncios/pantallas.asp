<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Pantallas</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->

        <%
            thisSystem = "anuncios"
            thisProcess = "anu.0120"
            SysLockOut
        %>          

        <style>
            .verde01 {
                color: #fff;
                background-color: rgb(61, 120, 65);
                border-color: rgb(61, 120, 65);
            }

            .verde02 {
                color: #fff;
                background-color: rgb(116, 179, 120);
                border-color: rgb(116, 179, 120);
            }

            .verde03 {
                color: #fff;
                background-color: rgb(175, 204, 177);
                border-color: rgb(175, 204, 177);
            }

            .celeste {
                color: #fff;
                background-color: rgb(71, 170, 206);
                border-color:  rgb(71, 170, 206);
            }            
        </style>   
    </head>

    <body plantilla="lista" reserva="175">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->  
        <%
            dim cc, t, tt, sqlString, data, labels, cuantas

            sqlString = "SELECT * FROM seg_anuncios_Lista_Pantallas;" 

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn") 
            set t = cc.execute(sqlString)        
        %>

        <div style="width 100%; margin: auto;">
            <br />

            <table style="width:95%; margin: auto;">
                <tr class="noborder">
                    <td colspan="3" style="font-size: 26px; text-align:left; width: 20%;">
                        &nbsp;Pantallas En Mi Red
                    </td>

                    <td style="text-align: right;">
                        <button type="button" class="form-btn verde "  onclick="crear()">
                            <i class=" fa fa-edit fa-xl" title="Nueva Pantalla"></i>
                        </button>                        
                    </td>                    
                </tr>
            </table>

            <table style="width:95%; margin: auto;">
                <tr style="font-size: 14px; background-color: rgb(61, 61, 61); color:rgb(255,255,255);">
                    <td style="padding: 5px; text-align:center; width: 10%;">Codigo</td>
                    <td style="padding: 5px; text-align:center; width: 25%;">Nombre</td>
                    <td style="padding: 5px; text-align:center; width: 40%;"> Descripcion</td>
                    <td style="padding: 5px; text-align:center; width: 25%;">&nbsp;</td>
                </tr>

                <tr>
                    <td colspan="5">
                        <div id="overFlow" style="width:100%; height: 500px; overflow: auto; background-color: rgb(207, 207, 207);" class="borde">
                            <table style="width: 100%;">
                                <%  
                                    cuantas = 0

                                    if not (t.bof or t.eof) then  
                                        Do                                  
                                            cuantas = cuantas + 1
                                %>

                                        <tr style="font-size: 14px; <%
                                                                        if Request.Cookies("Pantalla") = t("Pantalla") then
                                                                            response.write "background-color: rgb(213, 218, 237); "
                                                                        else
                                                                            response.write "background-color: rgb(255, 255, 255); "
                                                                        end if
                                                                    %> color:rgb(0,0,0); border-bottom: 1px solid rgb(194, 194, 194);">
                                            <td style="padding: 5px; text-align:center; width: 10%;" onclick="editar('<%= t("Pantalla") %>')"><%= t("Pantalla") %></td>
                                            <td style="padding: 5px; text-align:center; width: 25%;" onclick="editar('<%= t("Pantalla") %>')"><%= t("Nombre") %></td>
                                            <td style="padding: 5px; text-align:center; width: 40%;" onclick="editar('<%= t("Pantalla") %>')"><%= t("Descripcion") %></td>
                                            <td style="padding: 5px; text-align:center; width: 25%;">
                                                <button type="button" class="form-btn <%
                                                                        if t("Activas") > 0 then
                                                                            response.write "verde01"
                                                                        else
                                                                            if t("Publicaciones") > 0 then
                                                                                response.write "verde02"
                                                                            else
                                                                                response.write "verde03"
                                                                            end if
                                                                        end if
                                                                    %> " onclick="ver('<%= t("Pantalla") %>')">
                                                    <i class=" fa fa-list fa-xl" title="Editar Pantalla"></i>
                                                </button>

                                                <% if ParametroUsuario("anuncios", "anuncios_asgnar_usu") <> "" then %>
                                                    <button type="button" class="form-btn violeta " onclick="asignar('<%= t("Pantalla") %>')">
                                                        <i class=" fa fa-user fa-xl" title="Asignar Usuarios"></i>
                                                    </button>
                                                <% end if %>

                                                <button type="button" class="form-btn <%
                                                                        if Request.Cookies("Pantalla") = t("Pantalla") then
                                                                            response.write "azul"
                                                                        else
                                                                            response.write "celeste"
                                                                        end if
                                                                    %> " onclick="tv_default('<%= t("Pantalla") %>')">
                                                    <i class=" fa fa-tv fa-xl" title="Editar Pantalla"></i>
                                                </button>

                                                <button type="button" class="form-btn rojo " onclick="borrar('<%= t("Pantalla") %>')">
                                                    <i class=" fa fa-trash fa-xl" title="Borrar Pantalla"></i>
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
                    <td colspan="12" style="padding: 10px; text-align:center; width: 100%;">
                        &nbsp;&nbsp;Pantallas:&nbsp;<%= cuantas %>&nbsp;&nbsp;&nbsp;&nbsp;
                    </td>
                </tr>                               
            </table>
        </div>

        <%
            t.close: set t = nothing
            cc.close: set cc = nothing
        %>

        <script>
            function crear() {
                var vinculo ="editar_pantalla.asp?a=*";
                window.location.href = vinculo;
            }   

            function tv_default(codigo) {
                var vinculo ="asignar_tv.asp?a=" + codigo;
                window.location.href = vinculo;
            }   

            function editar(codigo) {
                var vinculo ="editar_pantalla.asp?a=" + codigo;
                window.location.href = vinculo;
            }   
           
            function asignar(pantalla) {
                var vinculo ="pantallas_asignar.asp?p=" + pantalla;
                window.location.href = vinculo;
            }              
          
            function ver(codigo) {
                var vinculo ="lista.asp?tv=" + codigo;
                window.location.href = vinculo;
            }             

            function borrar(codigo) {
                var confirmacion = confirm("Desea borrar la publicación seleccionada?");
                var vinculo ="borrar_pantalla.asp?a=" + codigo;

                if (confirmacion) {     
                    window.location.href = vinculo;
                };
            }    
        </script>     
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>