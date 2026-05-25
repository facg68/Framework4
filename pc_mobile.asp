<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <title><%= Application("Nombre") %></title>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>      
        <!-- #include virtual = "/core/includes/no_sql_injection.asp" -->          
    </head>

    <body onload="consultar()" style="overflow:hidden;">
        <script>
            function consultar() {
                if (esTelefono() == true) {
                    window.location.href = "/forma/mobile";
                } else {
                    window.location.href = "/forma/desktop";
                }
            }

            function esTelefono() {
                // 1. User Agent tradicional
                let ua = navigator.userAgent || navigator.vendor || window.opera;
                if (/android|iphone|ipad|ipod|blackberry|iemobile|opera mini/i.test(ua.toLowerCase())) {
                    return true;
                }

                // 2. User Agent moderno
                if (navigator.userAgentData && navigator.userAgentData.mobile !== undefined) {
                    return navigator.userAgentData.mobile;
                }

                /*
                // 3. Viewport real (más preciso que screen)
                const vw = Math.min(window.innerWidth, window.innerHeight);
                if (vw <= 800) {  // Ajustable a tu gusto
                    return true;
                }
                */

                return false;
            }
        </script>      
    </body>
</html>
