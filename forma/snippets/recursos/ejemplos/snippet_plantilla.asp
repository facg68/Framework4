<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    Snip_Width = 450 ' Tamaño de ejemplo en px
%>

<!-- #include virtual = "/core/includes/snippets.inc" -->

<!--
    PLANTILLA PARA LOS SNIPPETS 

    El "nomSnip" DEBE CORRESPONDER EXACTAMENTE con el nombre del Snippet 
    en la tabla de Procesos, de otro modo, el Snippet no se cargará

    ⚠️ OBLIGATORIO:
    Debe definirse Snip_Width y llamarse en redimWindow().
    Si no se hace, el Snippet NO se mostrará.

    Cada Snippet puede tener un tamaño distinto. Pero debe definirse ya que
    Desktop lo usa para manejar el espacio disponible

    Para facilitar el proceso, se puede reemplazar
    nomSnip -> nombre o diminutivo de este snippet
-->

<style>
    /*
        El nombre de las clases CSS deben 
        empezar con el nombre del snippet 
    */

    .nomSnip_clase {
    }
</style>

<%
    '
    ' Las funciones y procedimientos que usa el Snippet
    ' deben empezar con el nombre del Smippet
    '
    function nomSnip_LoQueSea()
    end function
%>

<div style="padding: 5px; background-color: rgb(235, 235, 235); max-width: <%= Snip_Width %>px;">  
    Programación del Snippet
</div>  

<script>
    /*
        Reglas para Snippets en JavaScript:

        01. Las funciones y procedimientos en JavaScript que use
            el Snippet deben empezar con el nombre del Snippet

        02. ⚠️ No hay variables globales.

            Si necesitas mantener estado o compartir información dentro
            del snippet, considera implementar funciones tipo _set() / _get().

            La forma de persistir los datos queda a tu criterio.

            Pista: El navegador ya ofrece mecanismos para persistir información.

        03. Se debe verificar que los objetos existan antes de usarlos,
            ya que pueden no estar disponibles en el DOM al momento
            de invocarlos.
    */

    function nomSnip_LoQueSea() {
    }

    function nomSnip_init() {
        /*
            Esta función se usa para verificar información o 
            inicializar o modificar el Snippet justo antes
            de presentarlo en el escritorio.

            Podemos redimensionar la ventana (pero no menor al 
            tamaño de los elementos internos) usando la función
            redimWindow.

            Esta función se ejecuta automáticamente cuando 
            se crea la ventana
        */

        redimWindow("nomSnip", <%= Snip_Width %>)
    }    

    function nomSnip_close() {
        /*
            Esta función es 100% opcional. Se usa para limpiar 
            información o cerrar cualquier conexión o guardar
            estados en tablas.

            Si no se va a hacer uso de estos servicios, 
            no es necesario tener esta función en el Snippet.

            Esta función se ejecuta automáticamente cuando 
            se cierra la ventana
        */
    }        
</script> 