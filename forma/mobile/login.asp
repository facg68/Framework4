<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <meta charset="UTF-8" />
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

            .login-screen {
                width: 100%;
                max-width: 420px;
                padding: 40px 28px;
                box-sizing: border-box;
            }

            .logo {
                text-align: center;
                font-size: 2.2rem;
                font-weight: 700;
                margin-bottom: 8px;
            }

            .subtitle {
                text-align: center;
                font-size: 1rem;
                opacity: 0.75;
                margin-bottom: 35px;
            }

            label {
                display: block;
                font-size: 0.9rem;
                margin-bottom: 6px;
                opacity: 0.8;
            }

            input, select {
                width: 100%;
                padding: 14px;
                margin-bottom: 18px;

                border: none;
                border-radius: 14px;

                background: rgba(255,255,255,0.08);
                color: white;
                font-size: 1rem;

                outline: none;
                box-sizing: border-box;
            }

            input:focus, select:focus {
                background: rgba(255,255,255,0.12);
                box-shadow: 0 0 0 2px rgba(70,138,190,0.6);
            }

            select option {
                background: #111;
                color: white;
            }

            .btn-login {
                width: 100%;
                padding: 15px;
                margin-top: 10px;

                border: none;
                border-radius: 16px;

                background: linear-gradient(135deg, #02576e, #468abe);
                color: white;

                font-size: 1.1rem;
                font-weight: bold;

                cursor: pointer;
                transition: transform 0.15s ease;
            }

            .btn-login:active {
                transform: scale(0.97);
            }

            .footer {
                margin-top: 40px;
                text-align: center;
                font-size: 0.8rem;
                opacity: 0.35;
            }
        </style>
    </head>

    <body>
        <%
            Session("c2_Reset") = ""
            Response.Cookies("c2_Reset") = "Mar 20 Feb 11:38"

            Response.Cookies("usuario").Expires = Now
            Response.Cookies("nombre").Expires = Now
            Response.Cookies("usuPath").Expires = Now
            Response.Cookies("max_WP").Expires = Now
        %>

        <div class="login-screen">
            <div class="logo">
                La Extranet de Fabrizio
            </div>

            <form action="login2.asp" method="post">
                <input id="esTelefono" name="esTelefono"
                    type="hidden"
                    value="<%= Request.QueryString("t") %>">

                <label>Usuario</label>
                <input id="txtUsuario" name="txtUsuario"
                    type="text" required>

                <label>Password</label>
                <input id="txtPassword" name="txtPassword"
                    type="password" required>

                <label>Conexión</label>
                <select id="chkMantener" name="chkMantener" required>
                    <option value="0">No Guardar</option>
                    <option value="1">Recordar</option>
                </select>

                <label>Menú</label>
                <select id="chkMenu" name="chkMenu" required>
                    <option value="0">No Crear</option>
                    <option value="1">Crear Menú</option>
                </select>

                <button class="btn-login" type="submit">
                    Entrar
                </button>
            </form>

            <div class="footer">
                <%= "Fabrizio Cárdenas, " & Year(Now()) %>
            </div>

        </div>

    </body>
</html>