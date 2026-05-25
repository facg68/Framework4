<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Editar Colecciones</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  

        <%
            thisSystem = "discos"
            thisProcess = "discos.0260"
            SysLockOut
        %>              
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    
        <%
            Usuario = Request.Cookies("usuario")
            ordenamiento = request.QueryString("o")
        %>  

        <br />

        <form name="form_transaccion" id="form_transaccion" method="post" action="grabar_nuevo.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 40%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Nueva Colección
                </div>
                
                <div style="flex: 0 0 60%; text-align: right;">
                    <button class="form-btn verde normal" type="button" onclick="grabar()">Grabar</button>
                    <button class="form-btn rojo normal"  type="button" onclick="volver()">Cancelar</button>
                </div>
            </div>        

            <div class="main main-scroll">
                <div class="no-ver">
                    <input type="text" id="Usuario" name="Usuario" value="<%= Usuario %>">
                    <input type="text" id="Orden"   name="Orden"   value="<%= Ordenamiento %>">                                
                </div>

                <div class="line">
                    <label class="label normal">Nombre</label>
                    <input class="field xl" type="text" id="Nombre" name="Nombre" required>
                </div>   

                <div class="line">
                    <label class="label normal">Descripcion</label>
                    <textarea class="field xxl" 
                                name="Descripcion" id="Descripcion" 
                                rows=3 cols=80></textarea>
                </div> 

                <div class="line">
                    <label class="label normal">Predeterminada</label>
                    <select class="field tiny" name="PorDefecto" id="PorDefecto" required >
                        <option value="1">Si</option>
                        <option value="0" selected>No</option>               
                    </select>                 
                </div>                    
            </div>                
        </form>

        <br />

        <script>
            function grabar() {
                document.getElementById("form_transaccion").submit(); 
            }    

            function volver() {
                var vinculo = "lista.asp?o=<%= ordenamiento %>";
                window.location.href = vinculo;
            }   
        </script>

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->    
    </body>
</html>