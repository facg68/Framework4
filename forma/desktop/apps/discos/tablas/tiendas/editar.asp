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
        %>       
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    
        <%
            dim con, t, tt, sqlString, cbox
            dim Usuario, Codigo, Nombre, Contacto, SitioWeb, Correo, Tipo, Pais
            dim Telefono1, Telefono2, Direccion, Notas, MediosDigitales, MediosFisicos
            dim Musica, Video, Juegos, Software, Libros, Hardware, Estatus

            Usuario = Request.Cookies("usuario")
            Codigo = Request.QueryString("c")
            Grupo = Request.QueryString("g")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

                sqlString = "SELECT Usuario, Codigo, Nombre, Contacto, SitioWeb, Correo, Tipo, Pais, Telefono1, Telefono2, Direccion, " & _
                            " Notas, MediosDigitales, MediosFisicos, Musica, Video, Juegos, Software, Libros, Hardware, Estatus " & _
                                "FROM  dbo.discos_Tiendas " & _
                                "WHERE (Usuario = '" & Usuario & "') " & _
                                "AND (Codigo = '" & Codigo & "');"

                set t = con.execute(sqlString)
                    Nombre = t("Nombre")
                    Contacto = t("Contacto")
                    SitioWeb = t("SitioWeb")
                    Correo = t("Correo")
                    Tipo = t("Tipo")
                    Pais = t("Pais")
                    Telefono1 = t("Telefono1")
                    Telefono2 = t("Telefono2")
                    Direccion = t("Direccion")
                    Notas = t("Notas")
                    MediosDigitales = t("MediosDigitales")
                    MediosFisicos = t("MediosFisicos")
                    Musica = t("Musica")
                    Video = t("Video")
                    Juegos = t("Juegos")
                    Software = t("Software")
                    Libros = t("Libros")
                    Hardware = t("Hardware")
                    Estatus = t("Estatus")
                t.close: set t = nothing
        %>  

        <br />

        <form name="form_transaccion" id="form_transaccion" method="post" action="grabar_registro.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    <%= Nombre %>
                </div>
                
                <div style="flex: 0 0 50%; text-align: right;">
                    <button class="form-btn verde normal" onclick="grabar()">Guardar</button>                   
                    <button class="form-btn rojo normal"  onclick="cancelar()">Cancelar</button>
                </div>
            </div>        

            <div class="main main-scroll">
                <div class="no-ver">
                    <input type="text" id="grupo" name="grupo" value="<%= grupo %>" required>
                    <input type="text" id="Codigo" name="Codigo" value="<%= Codigo %>" required>
                    <input type="text" id="Usuario" name="Usuario" value="<%= Usuario %>" required>
                </div>

                <div class="line">
                    <label class="label normal">Nombre</label>
                    <input class="field xl" type="text" id="Nombre" name="Nombre" required value="<%= Nombre %>">
                </div>   

                <div class="line">
                    <label class="label normal">Contacto</label>
                    <input class="field xl" type="text" id="Contacto" name="Contacto" required value="<%= Contacto %>">
                </div>     

                <div class="line">
                    <label class="label normal">Sitio Web</label>
                    <input class="field xl" type="text" id="SitioWeb" name="SitioWeb" required value="<%= SitioWeb %>">
                </div>   

                <div class="line">
                    <label class="label normal">Correo</label>
                    <input class="field xl" type="text" id="Correo" name="Correo" required value="<%= Correo %>">
                </div>                                                        

                <div class="line">
                    <label class="label normal">Tipo</label>
                    <select class="field large" name="Tipo" id="Tipo" required>
                        <option value="1" <% if Tipo = "1" then response.write " selected" %>>Tienda Fisica</option>
                        <option value="0" <% if Tipo = "0" then response.write " selected" %>>Tienda Online</option>             
                    </select>                 
                </div>   

                <div class="line">
                    <label class="label normal">País</label>
                    <%
                        set tt = con.execute("SELECT Codigo, Nombre FROM seg_Paises ORDER BY Nombre ASC;")
                            if not (tt.bof or tt.eof) then
                                response.write "<select class='field large' name='Pais' id='Pais' required >" 
                                    Do
                                        response.write "<option value='" & tt("Codigo") & "' "
                                            if Pais = tt("Codigo") then response.write " selected" 
                                        response.write ">" & tt("Nombre") & "</option>"

                                        tt.MoveNext
                                    Loop Until tt.eof
                                response.write "</select>"
                            end if
                        tt.close: set tt = nothing
                    %>                    
                </div>      

                <div class="line">
                    <label class="label normal">Teléfono 1</label>
                    <input class="field normal" type="text" id="Telefono1" name="Telefono1" required value="<%= Telefono1 %>">
                </div>   

                <div class="line">
                    <label class="label normal">Teléfono 2</label>
                    <input class="field normal" type="text" id="Telefono2" name="Telefono2" required value="<%= Telefono2 %>">
                </div>

                <div class="line">
                    <label class="label normal">Dirección</label>
                    <label class="label full">
                        <textarea class="field full" 
                                  name="Direccion" id="Direccion" 
                                  rows=3 cols=80><%= Direccion %></textarea>
                </div>

                <div class="line">
                    <label class="label normal">Notas</label>
                    <label class="label full">
                        <textarea class="field full" 
                                  name="Notas" id="Notas" 
                                  rows=3 cols=80><%= Notas %></textarea>
                </div>     

                <div class="line">
                    <label class="label normal">Medios Físicos</label>
                    <select class="field tiny" name="MediosFisicos" id="MediosFisicos" required >
                        <option value="1" <% if MediosFisicos = 1 then response.write " selected" %>>Si</option>
                        <option value="0" <% if MediosFisicos = 0 then response.write " selected" %>>No</option>               
                    </select>                 
                </div>

                <div class="line">
                    <label class="label normal">Medios Digitales</label>
                    <select class="field tiny" name="MediosDigitales" id="MediosDigitales" required >
                        <option value="1" <% if MediosDigitales = 1 then response.write " selected" %>>Si</option>
                        <option value="0" <% if MediosDigitales = 0 then response.write " selected" %>>No</option>               
                    </select>                 
                </div>                                

                <div class="line">
                    <label class="label normal">Medio Musical</label>
                    <select class="field tiny" name="Musica" id="Musica" required >
                        <option value="1" <% if Musica = 1 then response.write " selected" %>>Si</option>
                        <option value="0" <% if Musica = 0 then response.write " selected" %>>No</option>               
                    </select>                 
                </div>

                <div class="line">
                    <label class="label normal">Videos/Películas</label>
                    <select class="field tiny" name="Video" id="Video" required >
                        <option value="1" <% if Video = 1 then response.write " selected" %>>Si</option>
                        <option value="0" <% if Video = 0 then response.write " selected" %>>No</option>               
                    </select>                 
                </div>

                <div class="line">
                    <label class="label normal">Videojuegos</label>
                    <select class="field tiny" name="Juegos" id="Juegos" required >
                        <option value="1" <% if Juegos = 1 then response.write " selected" %>>Si</option>
                        <option value="0" <% if Juegos = 0 then response.write " selected" %>>No</option>               
                    </select>                 
                </div>

                <div class="line">
                    <label class="label normal">Software</label>
                    <select class="field tiny" name="Software" id="Software" required >
                        <option value="1" <% if Software = 1 then response.write " selected" %>>Si</option>
                        <option value="0" <% if Software = 0 then response.write " selected" %>>No</option>               
                    </select>                 
                </div>      

                <div class="line">
                    <label class="label normal">Libros</label>
                    <select class="field tiny" name="Libros" id="Libros" required >
                        <option value="1" <% if Libros = 1 then response.write " selected" %>>Si</option>
                        <option value="0" <% if Libros = 0 then response.write " selected" %>>No</option>               
                    </select>                 
                </div>  

                <div class="line">
                    <label class="label normal">Periférico</label>
                    <select class="field tiny" name="Hardware" id="Hardware" required >
                        <option value="1" <% if Hardware = 1 then response.write " selected" %>>Si</option>
                        <option value="0" <% if Hardware = 0 then response.write " selected" %>>No</option>               
                    </select>                 
                </div>

                <div class="line">
                    <label class="label normal">Estatus</label>
                    <select class="field large" name="Estatus" id="Estatus" required >
                        <option value="1" <% if Estatus = 1 then response.write " selected" %>>Este formato es Actual</option>
                        <option value="0" <% if Estatus = 0 then response.write " selected" %>>El formato ya es obsoleto</option>           
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

        <% con.close: set con = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->        
    </body>
</html>