<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Título de mi Página" %>
        <title><%= PageTitle %></title>
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <div class="page-title-bar">
            <div class="title"><%= PageTitle %></div>
        </div>

        <main>
            <div class="contenedor">
                <button onclick="alert('¡Hola desde ASP!')">
                    Presióname
                </button>
            <div>
        </main>

        <footer class="footer-contextual">
            <button class="footer-button" type="button" aria-label="Volver" onclick="volver()">
                <i class="fas fa-undo-alt"></i>
            </button>
        </footer>

        <script>
            function volver() {
                history.back();
            }
        </script>

        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>