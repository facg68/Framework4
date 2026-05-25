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
            dim con, t, tt, sqlString, cbox, ordenamiento
            dim Usuario, Codigo, Nombre, Descripcion, PorDefecto

            Usuario = Request.Cookies("usuario")
            Codigo = Request.QueryString("c")
            ordenamiento = request.QueryString("o")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            sqlString = "SELECT Usuario, Codigo, Nombre, Descripcion, PorDefecto, DeSistema " & _
                        "FROM dbo.discos_Carpetas " & _
                        "WHERE (Usuario = '" & Usuario & "') " & _
                        "AND (Codigo = '" & Codigo & "');"

            set t = con.execute(sqlString)
                Nombre = t("Nombre")
                Descripcion = t("Descripcion")
                PorDefecto = t("PorDefecto")
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
                    <input type="text" id="Codigo"  name="Codigo"  value="<%= Codigo %>">
                    <input type="text" id="Orden"   name="Orden"   value="<%= Ordenamiento %>">                                
                </div>

                <div class="line">
                    <label class="label normal">Nombre</label>
                    <input class="field xl" type="text" id="Nombre" name="Nombre" required value="<%= Nombre %>">
                </div>   

                <div class="line">
                    <label class="label normal">Descripcion</label>
                    <textarea class="field xxl" 
                                name="Descripcion" id="Descripcion" 
                                rows=3 cols=80><%= Descripcion %></textarea>
                </div> 

                <div class="line">
                    <label class="label normal">Predeterminada</label>
                    <select class="field tiny" name="PorDefecto" id="PorDefecto" required >
                        <option value="1" <% if PorDefecto = 1 then response.write " selected" %>>Si</option>
                        <option value="0" <% if PorDefecto = 0 then response.write " selected" %>>No</option>               
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

        <% con.close: set con = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->    
    </body>
</html>