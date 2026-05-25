<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Editar Foramto de Pantalla</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->   

        <style>
            body { overflow: auto;}
        </style>        
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->     

        <%
            dim con, t, tt, sqlString, cbox, ordenamiento
            dim Usuario, Codigo, Nombre, Musica, Video, Juegos, Software, Libros, Obsoleta

            Usuario = Request.Cookies("Usuario")
            Codigo = Request.QueryString("c")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            sqlString = "SELECT Nombre " & _
                        "FROM discos_FormatosPantalla " & _
                        "WHERE (Usuario = '" & Usuario & "') " & _
                        "AND (Codigo = '" & Codigo & "');"

            set t = con.execute(sqlString)
                Nombre = t("Nombre")
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
                    <input type="text" id="Usuario" name="Usuario" value="<%= Usuario %>">
                    <input type="text" id="Codigo" name="Codigo" value="<%= Codigo %>">
                </div>

                <div class="line">
                    <label class="label normal">Nombre</label>
                    <input class="field xl" type="text" id="Nombre" name="Nombre" value="<%= Nombre %>" required>
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

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->    
    </body>
</html>