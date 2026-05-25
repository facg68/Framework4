<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Editar Plataformas</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  

        <%
            thisSystem = "discos"
            thisProcess = "discos.0240"
            SysLockOut
        %>               
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    
        <%
            dim Usuario, Est, Tipo, Ordenamiento

            Usuario = Request.Cookies("usuario")
            Est = request.QueryString("e")
            Tipo = request.QueryString("t")
            ordenamiento = request.QueryString("o")
        %>  

        <br />

        <form name="form_transaccion" id="form_transaccion" method="post" action="grabar_nuevo.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 40%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Nueva Plataforma
                </div>
                
                <div style="flex: 0 0 60%; text-align: right;">
                    <button class="form-btn verde normal" type="button" onclick="grabar()">Grabar</button>
                    <button class="form-btn rojo normal"  type="button" onclick="volver()">Cancelar</button>
                </div>
            </div>        

            <div class="main main-scroll">
                <div class="no-ver">
                    <input type="text" id="Usuario"      name="Usuario"      value="<%= Usuario %>">
                    <input type="text" id="Estatus"      name="Estatus"      value="<%= Est %>">
                    <input type="text" id="Tipo"         name="Tipo"         value="<%= Tipo %>">
                    <input type="text" id="Ordenamiento" name="Ordenamiento" value="<%= Ordenamiento %>">                          
                </div>

                <div class="line">
                    <label class="label large">Nombre</label>
                    <input class="field xl" type="text" id="Nombre" name="Nombre" required>
                </div>           

                <div class="line">
                    <label class="label large">Es para Videojuegos</label>
                    <select class="field tiny" name="Juegos" id="Juegos" required >
                        <option value="1">Si</option>
                        <option value="0">No</option>
                    </select>                 
                </div>

                <div class="line">
                    <label class="label large">Es para Software</label>
                    <select class="field tiny" name="Software" id="Software" required >
                        <option value="1">Si</option>
                        <option value="0">No</option>           
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
                var vinculo = "lista.asp?e=<%= Est %>&t=<%= Tipo %>&o=<%= ordenamiento %>";
                window.location.href = vinculo;
            }   
        </script>
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->    
    </body>
</html>