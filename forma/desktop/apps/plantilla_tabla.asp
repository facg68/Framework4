<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <title>Ejemplo de Tablas</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->             
    </head>

    <body plantilla="tabla" tabla="95" reserva="185">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <br />

        <form> <!-- Si no es un Formulario, esto se omite! -->
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Ejemplo de Tablas
                </div>
                
                <div style="flex: 0 0 50%; text-align: right;">
                    <button class='form-btn rojo normal'  type='button' onclick="volver()">Cancelar</button>&nbsp;&nbsp;
                    <button class='form-btn verde normal' type='button' onclick="grabar()">Grabar</button>&nbsp;&nbsp;
                </div>
            </div>        

            <div class="main main-scroll" style="width: 95%;"> <!-- class='main' omite el scroll dentro del main -->
                <div class="line">
                    <div class="tabla-wrapper">
                        <table class="tabla tabla-green">
                            <thead>
                                <tr>
                                    <th class="sticky">Encabezado A</th>
                                    <th class="sticky">Encabezado B</th>
                                    <th class="sticky">Encabezado C</th>
                                </tr>
                            </thead>

                            <tbody>
                                <% for linea = 1 to 100 %>
                                        <tr>
                                            <td>Dato <%= linea %>A</td>
                                            <td>Dato <%= linea %>B</td>
                                            <td>Dato <%= linea %>C</td>
                                        </tr>
                                <% next %>                                                          
                            </tbody>

                            <tfoot>
                                <tr>
                                    <td class="sticky">Pie A</td>
                                    <td class="sticky">Pie B</td>
                                    <td class="sticky">Pie C</td>
                                </tr>
                            </tfoot>
                        </table>
                    </div>
                </div>

                <div class="line">
                    <label class="label normal">Ver Tabla Como</label>
                    <input class="field large" type="text">
                </div>
            </div>
        </form>
  
        <br /><br />   

        <script type="text/javascript">

        </script>

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>