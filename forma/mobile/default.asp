<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <style>
            .bienvenida {
                display: flex;
                flex-direction: column;
                line-height: 1.5;
            }

            .saludo {
                font-size: clamp(1.0em, 3.5vw, 1.5rem);
                color: #666;
            }

            .usuario {
                font-size: clamp(1.2rem, 4vw, 2rem);
                font-weight: 600;
                color: #111;
            }   
            
            .justificado {
                text-align: justify;
                text-justify: inter-word;
            }            
        </style>
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <main>
            <div class="contenedor">
                <div class="line bienvenida">
                    <span class="saludo">Bienvenido,</span>
                    <span class="usuario"><%= NombreUsuario() %></span>
                </div>

                <div class="line justificado">
                    Seleccione una opción del menú en la esquina 
                    superior izquierda.
                </div>

                <div class="line justificado">
                    Aqui encontrará los módulos de los sistemas a los 
                    que usted tenga acceso. No todas las aplicaciones 
                    tienen contrapartes telefónicas. 
                </div>

                <div class="line justificado">
                    Las aplicaciones sueler ser diferentes y hasta 
                    tener opciones distintas a las versiones de
                    escritorio.
                </div>                
            </div>
        </main>

        <footer>
            <p id="copyright">Derechos reservados, [FECHA AQUI], Nombre de mi Empresa.</p>

            <script>
                function Year() {
                    return new Date().getFullYear();
                }

                const parrafo = document.getElementById('copyright');
                parrafo.innerHTML = `&copy; ${Year()}  Fabrizio Arturo Cárdenas`;
            </script>
        </footer>

        <script>
            function toggleMenu() {
                document.getElementById('sidebar').classList.toggle('open');
            }
        </script>
    </body>
</html>