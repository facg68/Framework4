<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Editar Panfleto</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->

        <%
            thisSystem = "anuncios"
            thisProcess = "anu.0200"
            SysLockOut
        %>        
    </head>

    <body plantilla="normal" reserva="150">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            dim con, t, p, sqlString, codPanfleto, verPantalla, estatusPanfleto, ordenadoPor, noPantallas

            codPanfleto = request.querystring("p")
            estatusPanfleto = request.querystring("e")
            ordenadoPor = request.querystring("op")  

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            if codPanfleto <> "*" then
                set t = con.execute("SELECT * FROM seg_listaPanfletos WHERE Secuencia = " & codPanfleto & ";")
            end if

            Function Ya()
                Dim dia, mes, amo

                dia = Right("0" & Day(Now()), 2)
                mes = Right("0" & Month(Now()), 2)
                amo = Year(Now())

                Ya = dia & "/" & mes & "/" & amo & " 00:00"
            End Function
        %>

        <br />

        <div style="95%; margin: auto;">
            <form id="formulario"  name="formulario" method="post" action="grabar.asp">
                <input id="secuencia"       name="secuencia"        type="text" value="<%= codPanfleto %>"      class="no-ver" />
                <input id="estatusPanfleto" name="estatusPanfleto"  type="text" value="<%= estatusPanfleto %>"  class="no-ver" />
                <input id="ordenadoPor"     name="ordenadoPor"      type="text" value="<%= ordenadoPor %>"      class="no-ver" />

                <table style="width: 95%; margin: auto;"> 
                    <tr>
                        <td style="width: 55%; font-size: 24px;">
                            <h3><%
                                if codPanfleto = "*" then
                                    response.write "Nuevo Panfleto"
                                else
                                    response.write "Editar Panfleto"
                                end if
                                %>
                            </h3>
                        </td>

                        <td style="width: 45%; text-align: right;">
                            <button class="form-btn verde normal" type="submit">Grabar</button>       
                            
                            <a href="lista.asp?e=<%= estatusPanfleto %>&op=<%= ordenadoPor %>">
                                <button class="form-btn rojo normal" type='button'>Cancelar</button>
                            </a>                    
                        </td>
                    </tr>
                </table>

                <div class="main main-scroll">
                    <div class="line">
                        <label class="label normal">Titulo</label>
                        <input class=" field xxl" id="nombre" name="nombre" type="text" required <% if codPanfleto <> "*" then response.write "value='" & t("nombre") & "'" %> />
                    </div>

                    <div class="line">
                        <label class="label normal">Desde</label>
                        <input class="field normal"
                                id="Desde" name="Desde" type="text" placeholder="dd/mm/aaaa hh:mm"  
                                value='<% 
                                    if codPanfleto <> "*" then 
                                        response.write t("Desde2")
                                    else
                                        response.write Ya()
                                    end if 
                                %>' 
                                required
                        />
                    </div>

                    <div class="line">
                        <label class="label normal">Hasta</label>
                        <input class="field normal"
                                id="Hasta" name="Hasta" type="text" placeholder="dd/mm/aaaa hh:mm" 
                                value='<% 
                                    if codPanfleto <> "*" then 
                                        response.write t("Hasta2")
                                    end if
                                %>'
                                required 
                        />
                    </div>   
                </div>
            </form>
        </div>

        <script type="text/javascript">
            mask(document.getElementById('Desde'),    ['99/99/9999 99:99']);
            mask(document.getElementById('Hasta'),    ['99/99/9999 99:99']);
        </script>   

        <%
            if codPanfleto <> "*" then
                t.close: set t = nothing
            end if    

            con.close: set con = nothing
        %>   
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>