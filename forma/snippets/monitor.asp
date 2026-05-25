<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    monitor_Width = 1000 ' Tamaño de ejemplo en px
%>

<!-- #include virtual = "/core/includes/snippets.inc" -->

<style>
    .monitor_td {
        padding: 2px        !important;
        font-size: 14px     !important;
    }

    .monitor_main {
        max-width: 96%;
        margin: 0.5rem auto;
        padding: 0;
        background: transparent;
        border-radius: 0;
        box-shadow: none;
        display: flex;
        font-family: sans-serif;
    }    
    
    .snippet-btn {
        font-size: 12px;
        font-family: Verdana;
    }
</style>

<!-- TABLA CON DATOS -->
    <div class="monitor_main" style="max-height: 400px;">
        <div class="tabla-wrapper" id="monitor_body">
            <div style="text-align: center;">Cargando Datos...</div>
        </div>
    </div> 
<!-- FIN TABLA CON DATOS -->

<script>
    function monitor_init(){
        redimWindow("monitor", <%= monitor_Width %>);
        monitor_refresh();

        monitorInterval = setInterval(monitor_refresh, 5000);
    }    

    function monitor_close(){
        if (monitorInterval !== null) {
            clearInterval(monitorInterval);
            monitorInterval = null;
        }
    }

    function monitor_refresh() { 
        fetch("/forma/snippets/recursos/monitor_data.asp")
            .then(r => r.text())
            .then(html => {
                document.getElementById("monitor_body").innerHTML = html;
            });
    }

    function monitor_ver(idSnippet) {
        restoreWindow(idSnippet);
    }

    function monitor_min(idSnippet) {
        minimizeWindow(idSnippet);
    }

    function monitor_cerrar(idSnippet) {
        restoreWindow(idSnippet);
        closeWindow(idSnippet);
    }

    function monitor_verTodo() { win_global("restore"); }
    function monitor_minTodo() { win_global("min"); }
    function monitor_cerrarTodo() { win_global("close"); }
</script> 