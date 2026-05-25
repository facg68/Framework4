<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Ejemplo de formulario móvil" %>
        <title><%= PageTitle %></title>      

        <style>
            .btn.verde    { --btn-color: rgb(21, 140, 49) }
            .btn.rojo     { --btn-color: rgb(182, 68, 80); }

            .btn {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                padding: 0 0.75rem;
                font-family: 'Ruda';
                font-size: 1rem;
                line-height: 1;
                border-right: 0;
                border-radius: 0.5rem;
                color: rgb(255, 255, 255);
                background-color: var(--btn-color, transparent);
                border: 1px solid var(--btn-color);
                width: 40px;
                height: 40px;
            }

            .btn.small  { min-width:  60px; font-size: 1rem; color: rgb(255, 255, 255); }
            .btn.normal { min-width: 100px; font-size: 1rem; color: rgb(255, 255, 255); }
            .btn.large  { min-width: 200px; font-size: 1rem; color: rgb(255, 255, 255); }
        </style>  
    </head>

    <body reserva="125" plus="pPrueba1 : plus_nomPlus | lista ; 
                              pPrueba2 : plus_nomPlus | botones ; " > 
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <div class="page-title-bar">
            <div class="title"><%= PageTitle %></div>
        </div>

        <main>
            <br />
            
            <div class="contenedor">
                <div class="line">
                    <label>Nombre:</label>
                    <input type="text" id="nombre" name="nombre" />
                </div>

                <div class="line">
                    <label>Mensaje:</label>
                    <textarea id="mensaje" name="mensaje"></textarea>
                </div>

                <div class="line">
                    <button class="btn large verde" type="button" id="pPrueba1" >
                        Objeto Plus A                        
                    </button>                
                </div>

                <div class="line">
                    <button class="btn large rojo" type="button" id="pPrueba2" >
                        Objeto Plus B
                    </button>                
                </div>
            </div>
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

        <!-- #include virtual = "/forma/plus/plantilla_ejemplo.plus" --> 
        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->    
    </body>
</html>