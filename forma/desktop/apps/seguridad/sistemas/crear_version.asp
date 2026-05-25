<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Nueva Version</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        
        <%
            dim con, t, p, ptt, sqlString, Sistema, ordenadoPor
            dim nombre, Descripcion, ClaseApp, IndiceOrdenamiento, Icono, sBitacora

            thisSystem = "seguridad"
            thisProcess = "seg.0090"
            SysLockOut

            Sistema = request.querystring("s")
            ordenadoPor = request.querystring("o")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            function NombreSistema(Sistema)
                dim sqlString

                sqlString = "SELECT sysNombre FROM seg_Sistemas WHERE (sysCodigo ='" & Sistema & "');"

                set ptt = con.execute(sqlString)   
                    NombreSistema = ptt("sysNombre") 
                ptt.close: set ptt = nothing
            end function         

        %>        
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->   

        <div style="width: 98%; margin: auto;">
            <br />

            <table style="width: 95%; margin: auto;">
                <tr>
                    <td style="width: 60%; font-size: 24px;">
                        <%
                            response.write "Nueva Version para " & NombreSistema(Sistema)
                        %>
                    </td>

                    <td style="width: 40%; text-align: right;">
                        <button class='form-btn rojo normal'  type='button' onclick="volver()">Cancelar</button>&nbsp;&nbsp;
                        <button class='form-btn verde normal' type='button' onclick="grabar()">Grabar</button>&nbsp;&nbsp;
                    </td>
                </tr>
            </table>    

            <form id="formulario"  name="formulario" method="post" action="grabar_version_nueva.asp">
                <input id="ordenadoPor" name="ordenadoPor" type="text" value="<%= ordenadoPor %>" class="no-ver"/>
                <input id="Sistema" name="Sistema" type="text" value="<%= Sistema %>" class="no-ver"/>

                <div class="main main-scroll">
                    <div class="line">
                        <label class="label normal">Version</label>
                        <input class="field normal" id="Version" name="Version" type="text" required />
                    </div>

                    <div class="line">
                        <label class="label normal">Resumen</label>
                        <input class="field xl" id="Resumen" name="Resumen" type="text" required />
                    </div>

                    <div class="line">
                        <label class="label normal">Obligatoria *</label>
                        <select class="field normal" name="Obligatoria" id="Obligatoria">
                            <option value="0" selected >&nbsp;</option>
                            <option value="1" >Es Obligatoria</option>
                        </select>
                    </div>

                    <div class="line">
                        <span style="font-size: 14px; font-weight: notmal; text-align: left;">
                            * Las versiones sólo son obligatorias para las apps de escritorio
                        </span>
                    </div>
                </div>
            </form>
        </div>

        <script type="text/javascript">           
            function volver() {
                var vinculo = "versiones.asp?o=<%= ordenadoPor %>";
                window.location.href = vinculo;
            }

            function grabar() {
                document.getElementById("formulario").submit();
            }      
        </script> 

        <% con.close: set con = nothing %>
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->    
    </body>
</html>