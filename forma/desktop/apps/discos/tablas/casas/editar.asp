<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Editar Casas</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->   

        <%
            thisSystem = "discos"
            thisProcess = "discos.0220"
            SysLockOut
        %>          
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->     
        <%
            dim con, t, tt, sqlString, cbox, ordenamiento
            dim Usuario, Codigo, Nombre, Musica, Video, Juegos, Software, Libros, Obsoleta

            Usuario = Request.Cookies("usuario")
            Codigo = Request.QueryString("c")
            ver = request.QueryString("v")
            ordenamiento = request.QueryString("o")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

                sqlString = "SELECT Usuario, Codigo, Nombre, Musica, Video, Juegos, Software, Libros, Hardware, Obsoleta " & _
                            "FROM dbo.discos_Casas " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Codigo = '" & Codigo & "');"

                set t = con.execute(sqlString)
                    Nombre = t("Nombre")
                    Musica = t("Musica")
                    Video = t("Video")
                    Juegos = t("Juegos")
                    Software = t("Software")
                    Libros = t("Libros")
                    Hardware = t("Hardware")
                    Obsoleta = t("Obsoleta")
                t.close: set t = nothing
        %>  

        <br />

        <form name="form_transaccion" id="form_transaccion" method="post" action="grabar_registro.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 40%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    <%= Nombre %>
                </div>
                
                <div style="flex: 0 0 60%; text-align: right;">
                    <button class="form-btn verde normal" type="button" onclick="grabar()">Grabar</button>
                    <button class="form-btn rojo normal"  type="button" onclick="volver()">Cancelar</button>
                </div>
            </div>        

            <div class="main main-scroll">
                <div class="no-ver">
                    <input type="text" id="Codigo" name="Codigo" value="<%= Codigo %>">
                    <input type="text" id="Orden" name="Orden" value="<%= Ordenamiento %>">                                
                </div>

                <div class="line">
                    <label class="label normal">Nombre</label>
                    <input class="field xl" type="text" id="Nombre" name="Nombre" required value="<%= Nombre %>">
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
                    <label class="label normal">Obsoleta</label>
                    <select class="field small" name="Obsoleta" id="Obsoleta" required >
                        <option value="0" <% if Obsoleta = 0 then response.write " selected" %>>Activa</option>
                        <option value="1" <% if Obsoleta = 1 then response.write " selected" %>>Obsoleta</option>           
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
                var vinculo = "lista.asp?v=<%= ver %>&o=<%= ordenamiento %>";
                window.location.href = vinculo;
            }   
        </script>

        <% con.close: set con = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->    
    </body>
</html>