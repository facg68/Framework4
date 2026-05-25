<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">
		<!-- #include virtual = "/core/includes/no_sql_injection.asp" -->
		<!-- #include virtual = "/core/includes/menu_pass.inc" -->

        <title>Definir Clave de Acceso del Usuario</title>
        <%
            dim cc, sqlString

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")   		        
        %>    
    </head>

    <body style="background-color: rgb(242, 242, 242);">
        <div style="width: 95%; margin: auto;">
            <table style="width: 100%"> 
                <tr>
                    <td>     
                        <table style="width: 100%;">
                            <tr>
                                <td style="width: 75%;">
                                    <h3></i>No se pudo Establecer la Clave de Usuario</h3>
                                </td>
                                
                                <td style="width: 25%; text-align: right;">
                                    <button class='btn btn-danger'  type='button' onclick="volver()">Re-Intentar Proceso</button>&nbsp;&nbsp;
                                </td>
                            </tr>
                        </table>            

                        <div class="row mt">
                            <div class="col-lg-12">
                                <div class="form-panel">
                                    <div class="form">
                                        <form class="cmxform form-horizontal style-form" style="display: inline;">
                                            <div class="form-group ">
                                                <div class="col-lg-8">
                                                    <br />

                                                    Se encontraron errores en el proceso.<br /><br />

                                                    <%
                                                        usuario   = request.form("codigo")
                                                        nuevo_01 = limpiar(request.form("password_nuevo1"))
                                                        nuevo_02 = limpiar(request.form("password_nuevo2"))

                                                        if (nuevo_01 <> nuevo_02) then
                                                            response.write "Las Claves Nuevas no Coinciden. No se puede realizar el proceso."
                                                        else
                                                            sqlString = "exec dbo.seg_pa_ActualizarClaveUsuario '" & usuario & "','" & nuevo_01 & "'"
                                                            cc.execute(sqlString)

                                                            cc.execute("UPDATE seg_usuarios SET usuReset = 0 WHERE usuCodigo = '" & usuario & "';")

                                                            CrearMenu usuario
                                                            ActualizarSnippetsUsuario usuario

															Response.Cookies("usuario") = usuario
															Response.Cookies("nombre") = NombreUsuario(usuario)
															Response.Cookies("usuPath") = "/perfiles/" & usuario
															Response.Cookies("max_WP") = ContarWallpapers()
                                                            Response.Cookies("usu_WP") = wallPaperUsuario()

															Response.Cookies("usuario").Expires = Date() + 1
															Response.Cookies("nombre").Expires = Date() + 1
															Response.Cookies("usuPath").Expires = Date() + 1
															Response.Cookies("max_WP").Expires = Date() + 1
															Response.Cookies("usu_WP").Expires = Date() + 1

															Response.Redirect "/core"
                                                        end if
                                                    %>                                                     

                                                    <br /><br />
                                                    Realice el proceso nuevamente
                                                    <br />
                                                </div>
                                            </div>
                                        </form>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </td>
                </tr>
            </table>
        </div>
    </body>

    <script type="text/javascript">
        function volver() {
            var vinculo = "login.asp";
            window.location.href = vinculo;
        }
    </script>

    <% cc.close: set cc = nothing %>    
</html>