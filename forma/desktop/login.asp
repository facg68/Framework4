<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>La Extranet de Fabrizio</title>    
        
        <style>
            body{
                margin:0;
                color:#6a6f8c;
                background:#c8c8c8;
                background-image: url("/core/imagenes/login_fondo.jpg");                     
                font:600 16px/18px 'Open Sans',sans-serif;
            }

            *,:after,:before{box-sizing:border-box}
            .clearfix:after,.clearfix:before{content:'';display:table}
            .clearfix:after{clear:both;display:block}
            a{color:inherit;text-decoration:none}

            .login-wrap{
                width:100%;
                margin:auto;
                max-width:525px;
                min-height:670px;
                position:relative;
                background: url("/core/imagenes/bg.jpg") no-repeat center;
                box-shadow:0 12px 15px 0 rgba(0,0,0,.24),0 17px 50px 0 rgba(0,0,0,.19);
            }
            .login-html{
                width:100%;
                height:100%;
                position:absolute;
                padding:90px 70px 50px 70px;
                background:rgba(40,57,101,.9);
            }
            .login-html .sign-in-htm,
            .login-html .sign-up-htm{
                top:0;
                left:0;
                right:0;
                bottom:0;
                position:absolute;
                transform:rotateY(180deg);
                backface-visibility:hidden;
                transition:all .4s linear;
            }
            .login-html .sign-in,
            .login-html .sign-up,
            .login-form .group .check{
                display:none;
            }
            .login-html .tab,
            .login-form .group .label,
            .login-form .group .button{
                text-transform:uppercase;
            }
            .login-html .tab{
                font-size:22px;
                margin-right:15px;
                padding-bottom:5px;
                margin:0 15px 10px 0;
                display:inline-block;
                border-bottom:2px solid transparent;
            }
            .login-html .sign-in:checked + .tab,
            .login-html .sign-up:checked + .tab{
                color:#fff;
                border-color:#1161ee;
            }
            .login-form{
                min-height:345px;
                position:relative;
                perspective:1000px;
                transform-style:preserve-3d;
            }
            .login-form .group{
                margin-bottom:15px;
            }
            .login-form .group .label,
            .login-form .group .input,
            .login-form .group .button{
                width:100%;
                color:#fff;
                display:block;
            }
            .login-form .group .input,
            .login-form .group .button{
                border:none;
                padding:15px 20px;
                border-radius:25px;
                background:rgba(255,255,255,.1);
            }
            .login-form .group input[data-type="password"]{
                text-security:circle;
                -webkit-text-security:circle;
            }
            .login-form .group .label{
                color:#aaa;
                font-size:12px;
            }
            .login-form .group .button{
                background:#1161ee;
            }
            .login-form .group label .icon{
                width:15px;
                height:15px;
                border-radius:2px;
                position:relative;
                display:inline-block;
                background:rgba(255,255,255,.1);
            }
            .login-form .group label .icon:before,
            .login-form .group label .icon:after{
                content:'';
                width:10px;
                height:2px;
                background:#fff;
                position:absolute;
                transition:all .2s ease-in-out 0s;
            }
            .login-form .group label .icon:before{
                left:3px;
                width:5px;
                bottom:6px;
                transform:scale(0) rotate(0);
            }
            .login-form .group label .icon:after{
                top:6px;
                right:0;
                transform:scale(0) rotate(0);
            }
            .login-form .group .check:checked + label{
                color:#fff;
            }
            .login-form .group .check:checked + label .icon{
                background:#1161ee;
            }
            .login-form .group .check:checked + label .icon:before{
                transform:scale(1) rotate(45deg);
            }
            .login-form .group .check:checked + label .icon:after{
                transform:scale(1) rotate(-45deg);
            }
            .login-html .sign-in:checked + .tab + .sign-up + .tab + .login-form .sign-in-htm{
                transform:rotate(0);
            }
            .login-html .sign-up:checked + .tab + .login-form .sign-up-htm{
                transform:rotate(0);
            }

            .hr{
                height:2px;
                margin:60px 0 50px 0;
                background:rgba(255,255,255,.2);
            }
            .foot-lnk{
                text-align:center;
            }             
            
            .centered {
                position: fixed;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
            }                
        </style>
    </head>

    <body>
        <%        
            Session("c2_Reset") = ""
            Response.Cookies("c2_Reset") = "Mar 20 Feb 11:38"
        %>

        <form action="login2.asp" method="post">
            <div class="login-wrap centered">
                <div class="login-html">
                    <div style="color:rgb(99, 108, 138); font-size:24px; text-align:center;">
                        LA EXTRANET DE FABRIZIO<br /><br />
                        - framework 4.0 -
                    </div>
                    <br /><br /><br />
                    <input id="tab-1" type="radio" name="tab" class="sign-in" checked><label for="tab-1" class="tab">Iniciar Sesion</label>
                    <input id="tab-2" type="radio" name="tab" class="sign-up"><label for="tab-2" class="tab"></label> 
                    
                    <div class="login-form">
                        <div class="sign-in-htm">
                            <div class="group">
                                <label for="txtUsuario" class="label">Nombre de Usuario</label>
                                <input name="txtUsuario" id="txtUsuario" type="text" class="input" style="font-size: 18px;">
                            </div>
                            <div class="group">
                                <label for="txtPassword" class="label">Clave</label>
                                <input name="txtPassword" id="txtPassword" type="password" class="input" data-type="password" style="font-size: 18px;">
                            </div>
                            <div class="group">
                                <input name="chkMantener" id="chkMantener" type="checkbox" class="check" value="1">
                                <label for="chkMantener"><span class="icon"></span>&nbsp;&nbsp;Mantenerme Conectado</label>
                            </div>
                            <div class="group">
                                <input name="chkMenu" id="chkMenu" type="checkbox" class="check" value="1">
                                <label for="chkMenu"><span class="icon"></span>&nbsp;&nbsp;Recrear Estructura de Menu</label>
                            </div>                            
                            <div class="group">
                                <input type="submit" class="button" value="Entrar">
                            </div>
                        </div>
                    </div>
                </div>
            </div>        
        </form>
    </body>
</html>