<%@ CodePage=65001 %>
<% 
    '-------------------------------+
    ' Full Framework 4.0 - ASP Page |
    '-------------------------------+
%>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->   
        <%
            thisSystem = "CODIGO-SISTEMA"
            thisProcess = "CODIGO-PROCESO"
            SysLockOut
        %>               
    </head>

    <body>
        <!--
            Como se usa dentro de Framework 4 (Son ejemplos. Los valores pueden cambiar):
            --------------------------------------------------------------------------------------------------------------------

            Los valores por defecto son:

            plantilla   -> Normal
            reserva     -> 200
            grafica     -> 100%
            tabla       -> 100%

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

            Con "main-scroll" hacemos que el contenido del "main" flote.
            Si no queremos esto, entoces solo ponemos
            
            class='main'                   
        -->    
        
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <br />

        <form> <!-- Si no es un Formulario, esto se omite! -->
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Titulo de mi Pagina
                </div>
                
                <div style="flex: 0 0 50%; text-align: right;">
                    <!-- Aqui van los filtros y botones -->
                     
                    <button class='form-btn rojo normal'  type='button' onclick="volver()">Cancelar</button>&nbsp;&nbsp;
                    <button class='form-btn verde normal' type='button' onclick="grabar()">Grabar</button>&nbsp;&nbsp;
                </div>
            </div>        

            <div class="main main-scroll"> <!-- si se usa un <div class='main'> se omite el scroll dentro del "main" -->
                <!-- Aqui van los campos del formulario o los datos -->
            </div>
        </form>
  
        <br /><br />   

        <script type="text/javascript">
        </script>

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>