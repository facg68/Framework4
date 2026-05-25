<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Editar Pantalla</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->    

        <%
            thisSystem = "anuncios"
            thisProcess = "anu.0120"
            SysLockOut
        %>    
    </head>

    <body plantilla="normal" reserva="150">
      <!-- #include virtual = "/core/includes/kernel/body.inc" -->    

        <%
            dim con, t, p, sqlString, pantalla
            dim codPantalla, nomPantalla, desPantalla

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            pantalla = request.QueryString("a")

            if pantalla = "*" then
                codPantalla = "*"
                nomPantalla = ""
                desPantalla = ""
                estatus = 1
            else      
                set t = con.execute("SELECT * FROM seg_Anuncios_Pantallas WHERE Pantalla = '" & pantalla & "';")
                codPantalla = t("Pantalla")
                nomPantalla = t("Nombre")
                desPantalla = t("Descripcion")
                estatus = 0
                t.close: set t = nothing
            end if
        %>

        <div style="width: 95%; margin: auto;">
            <br/>

            <form id="formulario" name="formulario" method="post" action="grabar_pantalla.asp">
                <input id="estatus" name="estatus" type="text" value="<%= estatus %>" class="no-ver" />    

                <table style="width: 95%; margin: auto;"> 
                    <tr>
                        <td style="font-size: 25px; width: 55%;">
                            <%
                                if pantalla <> "*" then 
                                    response.write "Editar Pantalla"
                                else
                                    response.write "Crear Nueva Pantalla"
                                end if
                            %>
                        </td>

                        <td style="width: 55%; text-align: right;">
                            <button type="submit" class="form-btn verde normal">Grabar</button>                 

                            <a href='pantallas.asp'>
                                <button type='button' class="form-btn rojo normal">Cancelar</button>     
                            </a>
                        </td>
                    </tr>
                </table>

                <div class="main main-scroll">
                    <div class="line">
                        <label class="label normal">Codigo</label>

                        <% if pantalla = "*" then %>
                            <input class="field small id="codigo" name="codigo" type="text" value="" required />
                        <% else %>
                            <input id="codigo" name="codigo" type="text" value="<%= codPantalla %>" class="no-ver" />
                            <input class="field small" id="dispcodigo" name="dispcodigo" type="text" value="<%= codPantalla %>" disabled />
                        <% end if %>
                    </div>

                    <div class="line">
                        <label for="nombre" class="label normal">Nombre</label>
                        <input class="field xl" id="nombre" name="nombre" type="text" value="<%= nomPantalla %>" required />
                    </div>

                    <div class="line">
                        <label class="label normal">Descripcion</label>
                        <input class="field xxl" id="descripcion" name="descripcion" type="text" value="<%= desPantalla %>" required />
                    </div>
                </div>
            </form>
        </div>

        <% con.close: set con = nothing %> 
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
  </body>
</html>