<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<% 
    '===========================================+
    ' Framework 4.0 - Plantilla Base Oficial     |
    ' Autor: Fabrizio Arturo Cárdenas González   |
    ' Fecha: 11/02/2026                          |
    ' Uso: Páginas Desktop ASP Clásico           |
    '===========================================+
%>
<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->   
        <%
            'thisSystem = "CODIGO-SISTEMA"
            'thisProcess = "CODIGO-PROCESO"
            'SysLockOut

            Usuario = "f4_Test"
        %>

        <style>
            /* 
                Si NO queremos que el contenido del "main" flote dentro del contenedor, 
                (usando la clase "main main-scroll"), debemos cambiar el 
                overflow a "overflow: auto;" y definir el main como "main" sólo
             */

             .triggerPlus {
                font-family: Ruda;
                font-size: 18px;
             }
        </style>   
    </head>

    <body reserva="125" plus="pPrueba1 : plus_nomPlus @ mouseover | lista ; 
                              pPrueba2 : plus_nomPlus @ mouseover | botones ; 
                              pPrueba3 : plus_nomPlus @ mouseover | cinta ; " > 

        <!--
            Como se usa dentro de Framework 4 (Son ejemplos. Los valores pueden cambiar):
            --------------------------------------------------------------------------------------------------------------------

            Los valores por defecto son:

            plantilla   -> Normal
            reserva     -> 200
            grafica     -> 50%
            tabla       -> 50%

            <body>                                                              (se asume que la plantilla es "normal" y la reserva = 200)
            <body plantilla="normal" reserva="250">

            <body plantilla="tabla" tabla="95">                                 (se asume que la reserva = 200)
            <body plantilla="tabla" tabla="95" reserva="190">                   

            <body plantilla="grafica" grafica="95">                             (se asume que la reserva = 200)
            <body plantilla="grafica" grafica="95" reserva="190">      
            
            <body plantilla="dividida" tabla="60">                              (se asume que grafica = 40 y reserva = 200)
            <body plantilla="dividida" grafica="45">                            (se asume que tabla = 55 y reserva = 200)
            <body plantilla="dividida" tabla="60" grafica="40" reserva="190">

            <body plantilla="lista">                                            (se asume que la reserva = 200)
            <body plantilla="lista" reserva="190">

            <body plantilla="panoramica" grafica="95">                          (se asume que la reserva = 200)
            <body plantilla="panoramica" grafica="95" reserva="180">

            <body plantilla="area" grafica="95">                                (se asume que la reserva = 200)
            <body plantilla="area" grafica="95" reserva="180">              

            --------------------------------------------------------------------------------------------------------------------
        -->    

        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <br />

        <form> <!-- Si no es un Formulario, esto se omite! -->
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Titulo de mi Pagina
                    &nbsp;&nbsp;
                    <span class="triggerPlus" id="pPrueba1" style="cursor: pointer;"> [ +Lista ]</span>
                    &nbsp;&nbsp;
                    <span class="triggerPlus" id="pPrueba2" style="cursor: pointer;"> [ +Botones ]</span>
                    &nbsp;&nbsp;
                    <span class="triggerPlus" id="pPrueba3" style="cursor: pointer;"> [ +Cinta ]</span>
                </div>
                
                <div style="flex: 0 0 50%; text-align: right;">

                    <!-- ::::::::::::::::::::::::::::::::: -->
                    <!--  Controles de Filtros y Botones   -->
                    <!-- ::::::::::::::::::::::::::::::::: -->

                <button class='form-btn verde normal' type='button' onclick="guardar()">Guardar</button>&nbsp;&nbsp;
                <button class='form-btn azul normal'  type='button' onclick="cargar()">Cargar</button>&nbsp;&nbsp;                
                <button class='form-btn rojo normal'  type='button' onclick="borrar()">Borrar</button>&nbsp;&nbsp;         
                </div>
            </div>        

            <!-- ::::::::::::::::: -->
            <!--     Formulario    -->
            <!-- ::::::::::::::::: -->               

            <div class="main main-scroll"> 
                <!--
                    Con "main-scroll" hacemos que el contenido del "main" flote
                    dentro del contenedor. 

                    Si no queremos esto, entoces solo ponemos
                    
                    class='main'                   
                 -->

                <!-- ::::::::::::::::::::::::::::::::: -->
                <!--     Ejemplo de Uso de Máscaras    -->
                <!-- ::::::::::::::::::::::::::::::::: -->            

                <div class="line">
                    <label class="label small">Cédula</label>
                    <input class="field tiny" type="text" id="cedula" placeholder="8-888-888">
                </div>

                <div class="line">
                    <label class="label small">Teléfono</label>
                    <input class="field small" type="text" id="telefono" placeholder="6612-3456">
                </div>

                <div class="line">
                    <label class="label small">Fecha</label>
                    <input class="field tiny" type="text" id="fecha" placeholder="dd/mm/aaaa">
                </div>

                <div class="line">
                    <label class="label small">Hora</label>
                    <input class="field tiny" type="text" id="hora" placeholder="HH:MM">
                </div>

                <div class="line">
                    <label class="label small">Placa</label>
                    <input class="field tiny" type="text" id="placa" placeholder="AB1234">
                </div>

                <div class="line">
                    <label class="label small">VISA / MC</label>
                    <input class="field small" type="text" id="tarjeta" placeholder="9999 9999 9999 9999">
                </div>

                <div class="line">
                    <label class="label small">Mensaje:</label>
                    <textarea class="field xxl" type="text" id="mensaje" name="mensaje"></textarea>
                </div>


                <!--  
                    se puede usar un "label" especial como separador 
                    ó como "seccion": 
                -->

                <div class="line label-top">
                    <label class="label small">Grupo:</label>
                    <div class="label full section">
                        <div class="line-group">
                            <div class="line">
                                <label class="label small">Nombre:</label>
                                <input class="field normal" type="text" id="nombre" name="nombre" />
                            </div>

                            <div class="line">
                                <label class="label small">Mensaje 2:</label>
                                <textarea class="field xxl" type="text" id="mensaje2" name="mensaje2"></textarea>
                            </div>
                        </div>
                    </div>
                </div>  

                <div class="line label-top">
                    <label class="label small">Tabla:</label>
                    <div class="label full section">
                        <!-- ::::::::::::::::::::::::::::::: -->
                        <!--    Tabla Moderna Encapsulada    -->
                        <!-- ::::::::::::::::::::::::::::::: -->   

                        <table class="tabla tabla-blue" style="width: 98%; margin: auto;">
                            <thead>
                                <tr>
                                    <th>Nombre</th>
                                    <th>Correo</th>
                                    <th>Teléfono</th>
                                    <th class="sticky">Estado</th>
                                </tr>
                            </thead>

                            <tbody>
                                <tr>
                                    <td>Jorge Gomez</td>
                                    <td>jgomez@yahoo.com</td>
                                    <td>+507 1234-5678</td>
                                    <td>Activo</td>
                                </tr>

                                <tr>
                                    <td>Jorge Gomez</td>
                                    <td>jgomez@yahoo.com</td>
                                    <td>+507 1234-5678</td>
                                    <td>Activo</td>
                                </tr>

                                <tr>
                                    <td>Jorge Gomez</td>
                                    <td>jgomez@yahoo.com</td>
                                    <td>+507 1234-5678</td>
                                    <td>Activo</td>
                                </tr>                                

                                <!-- Más filas... -->
                            </tbody>
                        </table>     
                    </div>
                </div>                          
            </div>    
        </form>

        <br /><br />   

        <script type="text/javascript">
            const usuario = "<%= Usuario %>";
            const campos = ["cedula", "telefono", "fecha", "hora", "placa", "tarjeta", "mensaje", "nombre", "mensaje2"];

            function guardar() {
                campos.forEach(id => {
                    const input = document.getElementById(id);
                    F4Storage.set(usuario, id, input.value);
                    input.value = "";
                });
            }

            function cargar() {
                campos.forEach(id => {
                    var campo = document.getElementById(id);
                    campo.value = F4Storage.get(usuario, id);
                });
            }   
            
            function borrar() {
                campos.forEach(id => {
                    F4Storage.remove(usuario, id);
                });                
            }     
            
            /* 
               ::::::::::::::::::::::::::::::::::::::::::
               ::      Aplicación de las máscaras      ::
               ::::::::::::::::::::::::::::::::::::::::::
            */

            mask(document.getElementById('cedula'),     ['9-999-999', '99-999-999', 'AA09-999-9999']);
            mask(document.getElementById('telefono'),   ['9999-9999']);
            mask(document.getElementById('fecha'),      ['99/99/9999']);
            mask(document.getElementById('hora'),       ['99:99']);
            mask(document.getElementById('placa'),      ['AA9999']);
            mask(document.getElementById('tarjeta'),    ['9999 9999 9999 9999']);
        </script>

        <!-- #include virtual = "/forma/plus/ejemplos/plantilla_ejemplo.plus" --> 
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>