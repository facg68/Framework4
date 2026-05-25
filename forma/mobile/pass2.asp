<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <meta charset="UTF-8" />

        <!-- Seguridad -->
        <!-- #include virtual = "/core/includes/no_sql_injection.asp" -->
        <!-- #include virtual = "/core/includes/menu_pass.inc" -->

        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title><%= Application("Nombre") %></title>

        <style>
            body {
                margin: 0;
                padding: 0;
                min-height: 100dvh;

                display: flex;
                justify-content: center;
                align-items: center;

                font-family: system-ui, sans-serif;
                background: radial-gradient(circle at top, #1c1f26, #0b0d11);
                color: white;
            }

            .screen {
                width: 100%;
                max-width: 460px;
                padding: 45px 28px;
                box-sizing: border-box;
                text-align: center;
            }

            .title {
                font-size: 1.7rem;
                font-weight: 800;
                margin-bottom: 10px;
                color: #ffffff;
            }

            .subtitle {
                font-size: 1rem;
                opacity: 0.75;
                margin-bottom: 30px;
            }

            .message-box {
                padding: 0;
                margin-top: 20px;
                font-size: 1rem;
                line-height: 1.6;
                color: rgba(255,255,255,0.82);
            }

            .error-detail {
                margin-top: 18px;
                padding-left: 14px;
                border-left: 3px solid rgba(255,92,92,0.8);

                font-weight: 600;
                color: #ffffff;
            }

            .btn {
                margin-top: 35px;
                width: 100%;
                padding: 15px;

                border: none;
                border-radius: 16px;

                background: linear-gradient(135deg, #02576e, #468abe);
                color: white;

                font-size: 1.05rem;
                font-weight: bold;

                cursor: pointer;
                transition: transform 0.15s ease;
            }

            .btn:active {
                transform: scale(0.97);
            }

            .footer {
                margin-top: 40px;
                font-size: 0.8rem;
                opacity: 0.35;
            }
        </style>

        <%
            dim cc 

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")        
        %>
    </head>

    <body>
        <div class="screen">
            <div class="title">
                No se pudo establecer la clave
            </div>

            <div class="subtitle">
                Se encontraron errores en el proceso
            </div>

            <br />

            <div class="message-box">
                <%
                    usuario   = request.form("codigo")
                    nuevo_01  = limpiar(request.form("password_nuevo1"))
                    nuevo_02  = limpiar(request.form("password_nuevo2"))

                    if (nuevo_01 <> nuevo_02) then
                        %>
                            <div class="error-detail">
                                Las claves nuevas no coinciden.<br>
                                No se puede completar el proceso.
                            </div>
                        <%
                    else
                        sqlString = "exec dbo.seg_pa_ActualizarClaveUsuario '" & usuario & "','" & nuevo_01 & "'"

                        cc.execute(sqlString)

                        cc.execute("UPDATE seg_usuarios SET usuReset = 0 WHERE usuCodigo = '" & usuario & "';")

                        CrearMenu usuario

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

                        Response.Redirect "/forma/mobile/"
                    end if
                %>

                <br><br>
                Por favor, realice el proceso nuevamente.
            </div>

            <br />

            <button class="btn" onclick="window.location.href='/forma/mobile/login.asp'">
                Volver al Login
            </button>

            <div class="footer">

            </div>
        </div>
        <% cc.close: set cc = nothing %>
    </body>
</html>