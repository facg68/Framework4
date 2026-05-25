<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>

        <title>Editar Tienda</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "discos"
            thisProcess = "discos.0210"
            SysLockOut

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")            
        %>       
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    

        <br />

        <form name="form_transaccion" id="form_transaccion" method="post" action="grabar_nuevo.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Nueva Tienda
                </div>
                
                <div style="flex: 0 0 50%; text-align: right;">
                    <button class="form-btn verde normal" onclick="grabar()">Guardar</button>                   
                    <button class="form-btn rojo normal"  onclick="cancelar()">Cancelar</button>
                </div>
            </div>        

            <div class="main main-scroll">
                <div class="no-ver">
                    <input type="text" id="grupo" name="grupo" value="<%= grupo %>" required>
                    <input type="text" id="Usuario" name="Usuario" value="<%= Usuario %>" required>
                </div>

                <div class="line">
                    <label class="label normal">Nombre</label>
                    <input class="field xl" type="text" id="Nombre" name="Nombre" required>
                </div>   

                <div class="line">
                    <label class="label normal">Contacto</label>
                    <input class="field xl" type="text" id="Contacto" name="Contacto" required>
                </div>     

                <div class="line">
                    <label class="label normal">Sitio Web</label>
                    <input class="field xl" type="text" id="SitioWeb" name="SitioWeb" required >
                </div>   

                <div class="line">
                    <label class="label normal">Correo</label>
                    <input class="field xl" type="text" id="Correo" name="Correo" required>
                </div>                                                        

                <div class="line">
                    <label class="label normal">Tipo</label>
                    <select class="field large" name="Tipo" id="Tipo" required>
                        <option value="1">Tienda Fisica</option>
                        <option value="0">Tienda Online</option>             
                    </select>                 
                </div>   

                <div class="line">
                    <label class="label normal">País</label>
                    <%
                        set tt = con.execute("SELECT Codigo, Nombre FROM seg_Paises ORDER BY Nombre ASC;")
                            if not (tt.bof or tt.eof) then
                                response.write "<select class='field large' name='Pais' id='Pais' required >" 
                                    Do
                                        response.write "<option value='" & tt("Codigo") & "'>" & tt("Nombre") & "</option>"
                                        tt.MoveNext
                                    Loop Until tt.eof
                                response.write "</select>"
                            end if
                        tt.close: set tt = nothing
                    %>                    
                </div>      

                <div class="line">
                    <label class="label normal">Teléfono 1</label>
                    <input class="field normal" type="text" id="Telefono1" name="Telefono1">
                </div>   

                <div class="line">
                    <label class="label normal">Teléfono 2</label>
                    <input class="field normal" type="text" id="Telefono2" name="Telefono2">
                </div>

                <div class="line">
                    <label class="label normal">Dirección</label>
                    <label class="label full">
                        <textarea class="field full" 
                                  name="Direccion" id="Direccion" 
                                  rows=3 cols=80></textarea>
                </div>

                <div class="line">
                    <label class="label normal">Notas</label>
                    <label class="label full">
                        <textarea class="field full" 
                                  name="Notas" id="Notas" 
                                  rows=3 cols=80></textarea>
                </div>     

                <div class="line">
                    <label class="label normal">Medios Físicos</label>
                    <select class="field tiny" name="MediosFisicos" id="MediosFisicos" required >
                        <option value="1">Si</option>
                        <option value="0">No</option>               
                    </select>                 
                </div>

                <div class="line">
                    <label class="label normal">Medios Digitales</label>
                    <select class="field tiny" name="MediosDigitales" id="MediosDigitales" required >
                        <option value="1">Si</option>
                        <option value="0">No</option>               
                    </select>                 
                </div>                                

                <div class="line">
                    <label class="label normal">Medio Musical</label>
                    <select class="field tiny" name="Musica" id="Musica" required >
                        <option value="1">Si</option>
                        <option value="0">No</option>               
                    </select>                 
                </div>

                <div class="line">
                    <label class="label normal">Videos/Películas</label>
                    <select class="field tiny" name="Video" id="Video" required >
                        <option value="1">Si</option>
                        <option value="0">No</option>               
                    </select>                 
                </div>

                <div class="line">
                    <label class="label normal">Videojuegos</label>
                    <select class="field tiny" name="Juegos" id="Juegos" required >
                        <option value="1">Si</option>
                        <option value="0">No</option>               
                    </select>                 
                </div>

                <div class="line">
                    <label class="label normal">Software</label>
                    <select class="field tiny" name="Software" id="Software" required >
                        <option value="1">Si</option>
                        <option value="0">No</option>               
                    </select>                 
                </div>      

                <div class="line">
                    <label class="label normal">Libros</label>
                    <select class="field tiny" name="Libros" id="Libros" required >
                        <option value="1">Si</option>
                        <option value="0">No</option>               
                    </select>                 
                </div>  

                <div class="line">
                    <label class="label normal">Periférico</label>
                    <select class="field tiny" name="Hardware" id="Hardware" required >
                        <option value="1">Si</option>
                        <option value="0">No</option>               
                    </select>                 
                </div>    
            </div>
        </form>

        <br /><br />

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
        <% con.close: set con = nothing %>               
    </body>
</html>