<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Nueva Forma</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->   
        <%
            thisSystem = "discos"
            thisProcess = "discos.0205"
            SysLockOut
        %>  
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->     

        <%
            dim grupo
            grupo = request.QueryString("g")
        %>

        <br />

        <form name="form_transaccion" id="form_transaccion" method="post" action="grabar_nuevo.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Nueva Forma
                </div>
                
                <div style="flex: 0 0 50%; text-align: right;">
                    <button class="form-btn verde normal" onclick="grabar()">Guardar</button>                   
                    <button class="form-btn rojo normal"  onclick="cancelar()">Cancelar</button>
                </div>
            </div>        

            <div class="main main-scroll">
                <input class="no-ver" type="text" id="grupo" name="grupo" value="<%= grupo %>" required>

                <div class="line">
                    <label class="label normal">Nombre</label>
                    <input class="field xl" type="text" id="Nombre" name="Nombre" required>
                </div>            

                <div class="line">
                    <label class="label normal">Multi Lado</label>
                    <select class="field large" name="Multilados" id="Multilados" required >
                        <option value="1">El Medio Tiene 2 Lados</option>
                        <option value="0" selected>El Medio Tiene Solo 1 Lado</option>             
                    </select>                 
                </div>            

                <div class="line">
                    <label class="label normal">Medio Musical</label>
                    <select class="field tiny" name="Musica" id="Musica" required >
                        <option value="1" selected>Si</option>
                        <option value="0">No</option>               
                    </select>                 
                </div>

                <div class="line">
                    <label class="label normal">Videos/Películas</label>
                    <select class="field tiny" name="Video" id="Video" required >
                        <option value="1" selected>Si</option>
                        <option value="0">No</option>               
                    </select>                 
                </div>

                <div class="line">
                    <label class="label normal">Videojuegos</label>
                    <select class="field tiny" name="Juegos" id="Juegos" required >
                        <option value="1" selected>Si</option>
                        <option value="0">No</option>               
                    </select>                 
                </div>

                <div class="line">
                    <label class="label normal">Software</label>
                    <select class="field tiny" name="Software" id="Software" required >
                        <option value="1" selected>Si</option>
                        <option value="0">No</option>               
                    </select>                 
                </div>      

                <div class="line">
                    <label class="label normal">Libros</label>
                    <select class="field tiny" name="Libros" id="Libros" required >
                        <option value="1" selected>Si</option>
                        <option value="0">No</option>               
                    </select>                 
                </div>  

                <div class="line">
                    <label class="label normal">Periférico</label>
                    <select class="field tiny" name="Hardware" id="Hardware" required >
                        <option value="1" selected>Si</option>
                        <option value="0">No</option>               
                    </select>                 
                </div>

                <div class="line">
                    <label class="label normal">Estatus</label>
                    <select class="field large" name="Estatus" id="Estatus" required >
                        <option value="1" selected>Este formato es Actual</option>
                        <option value="0">El formato ya es obsoleto</option>           
                    </select>                 
                </div>
            </div>
        </form>

        <br />

        <script>
            function grabar() {
                document.getElementById("form_transaccion").submit(); 
            }    

            function cancelar() {
                var vinculo = "lista.asp?g=<%= grupo %>";
                window.location.href = vinculo;
            }
        </script>

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->        
    </body>
</html>