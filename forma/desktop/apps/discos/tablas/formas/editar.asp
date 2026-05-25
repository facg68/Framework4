<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Editar Formas</title>
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
            dim con, t, tt, sqlString, cbox
            dim Usuario, Forma, Nombre, Multilados, Musica, Video, Juegos, Software, Libros, Hardware, Icono_Forma, Estatus

            Usuario = Request.Cookies("usuario")
            Forma = Request.QueryString("c")
            grupo = request.QueryString("g")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

                sqlString = "SELECT Usuario, Forma, Nombre, Multilados, Musica, Video, Juegos, Software, Libros, Hardware, Icono_Forma, Estatus " & _
                            "FROM dbo.discos_Formas " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Forma = '" & Forma & "');"

                set t = con.execute(sqlString)
                    Nombre = t("Nombre")
                    Multilados = t("Multilados")
                    Musica = t("Musica")
                    Video = t("Video")
                    Juegos = t("Juegos")
                    Software = t("Software")
                    Libros = t("Libros")
                    Hardware = t("Hardware")
                    Icono_Forma = t("Icono_Forma")
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
                    <input type="text" id="Forma" name="Forma" value="<%= Forma %>" required>
                    <input type="text" id="Usuario" name="Usuario" value="<%= Usuario %>" required>
                </div>

                <div class="line">
                    <label class="label normal">Nombre</label>
                    <input class="field xl" type="text" id="Nombre" name="Nombre" required value="<%= Nombre %>">
                </div>            

                <div class="line">
                    <label class="label normal">Multi Lado</label>
                    <select class="field large" name="Multilados" id="Multilados" required>
                        <option value="1" <% if Multilados = 1 then response.write " selected" %>>El Medio Tiene 2 Lados</option>
                        <option value="0" <% if Multilados = 0 then response.write " selected" %>>El Medio Tiene Solo 1 Lado</option>             
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

        <br />

        <script>
            function guardar() {
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