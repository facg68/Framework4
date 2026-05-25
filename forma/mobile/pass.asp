<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>Definir Clave · <%= Application("Nombre") %></title>

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
                max-width: 420px;
                padding: 42px 28px;
                box-sizing: border-box;
            }

            .title {
                text-align: center;
                font-size: 1.8rem;
                font-weight: 700;
                margin-bottom: 6px;
            }

            .subtitle {
                text-align: center;
                font-size: 1rem;
                opacity: 0.7;
                margin-bottom: 35px;
            }

            label {
                display: block;
                font-size: 0.9rem;
                margin-bottom: 6px;
                opacity: 0.8;
            }

            input {
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

            input:focus {
                background: rgba(255,255,255,0.12);
                box-shadow: 0 0 0 2px rgba(70,138,190,0.6);
            }

            input[disabled] {
                opacity: 0.55;
                cursor: not-allowed;
            }

            .btn {
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

            .btn:active {
                transform: scale(0.97);
            }

            .footer {
                margin-top: 40px;
                text-align: center;
                font-size: 0.8rem;
                opacity: 0.35;
            }
        </style>

        <%
            '
            ' Init()
            '
            dim cc, tt, usuario

            if (session("c2_Reset") = "") OR (Request.Cookies("c2_Reset") = "") then
                response.redirect "/forma/mobile/login.asp"
            else
                if (session("c2_Reset") <> Request.Cookies("c2_Reset")) then
                    response.redirect "/forma/mobile/login.asp"
                else
                    if Cierres(session("c2_Reset"), Request.Cookies("c2_Reset")) <> 1 then
                        response.redirect "login.asp"
                    end if
                end if
            end if

            Session("c2_Reset") = ""
            Response.Cookies("c2_Reset") = ""

            function NombreUsuario()
                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                    set tt = cc.execute("SELECT usuNombre FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("rUsuario") & "';")
                        NombreUsuario = tt("usuNombre")
                    tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function

            function cierres(Cadena1, Cadena2)
                cierres = 1
                if left(Cadena1, 2) <> "/*" then cierres = 0
                if left(Cadena2, 2) <> "/*" then cierres = 0
                if right(Cadena1, 2) <> "$|" then cierres = 0
                if right(Cadena2, 2) <> "$|" then cierres = 0
            end function
        %>
    </head>

    <body>
        <div class="screen">

            <div class="title">
                Definir Clave
            </div>

            <div class="subtitle">
                Establece tu nueva contraseña
            </div>

            <form action="pass2.asp" method="post">
                <input type="hidden" name="esTelefono"
                    value="<%= Request.QueryString("t") %>">

                <input type="hidden" name="codigo"
                    value="<%= Request.Cookies("rUsuario") %>">

                <label>Usuario</label>
                <input type="text"
                    value="<%= NombreUsuario() %>"
                    disabled>

                <label>Nueva Clave</label>
                <input name="password_nuevo1"
                    id="password_nuevo1"
                    type="password"
                    required>

                <label>Repetir Clave</label>
                <input name="password_nuevo2"
                    id="password_nuevo2"
                    type="password"
                    required>

                <button class="btn" type="submit">
                    Guardar Clave
                </button>
            </form>

            <div class="footer">
                
            </div>

        </div>
    </body>
</html>