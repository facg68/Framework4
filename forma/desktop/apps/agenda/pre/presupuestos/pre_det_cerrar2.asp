<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>

        <title>Asignar Saldos</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->        

        <%
            dim con, t, sqlString, usu, pre
            dim contador, lblContador
            
            usu = Request.Cookies("usuario")
            pre = request.QueryString("p")

            '
            ' Abrimos la conexion con los datos...
            '

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")    
        %>

        <style>
            tr:not(:last-child) { border: none !important; }
        </style>      
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <div style="width:95%; margin: auto;">
            <br />

            <table style="width: 100%;">
                <tr>
                    <td style="width: 45%; text-align: left; font-size: 24px;">
                        Cerrar Presupuestos
                    </td>

                    <td style="width: 55%; text-align: right;">
                        <button class="form-btn naranja large" type="button" onclick = "cerrar()">Cerrar Presupuesto</button>&nbsp;

                        <a href="../lista.asp">
                            <button type="button" class="form-btn verde large">Cancelar Cierre</button>
                        </a>                    
                    </td>
                </tr>
            </table>

            <div class="main main-scroll">
                <div class="line">
                    <div class="full section" 
                         style="background-color: rgb(245, 245, 245); line-height: 1.6; font-family: 'Ruda Bold'; font-size: 17px; text-align: left;">
                        El Presupuesto que está por cerrar <u>tiene saldos pendientes</u> (no iguales a cero), por lo que el sistema generará 
                        registros de cierre en el presupuesto a actual, tanto para la cuenta de <span style="color: rgb(182, 68, 80);">CARTERA</span> y la cuenta de 
                        <span style="color: rgb(182, 68, 80);">EFECTIVO</span>, así como 
                        registros de apertura <span style="color: rgb(182, 68, 80);">PARA AMBAS CUENTAS</span> en el presupuesto que se seleccione de la lista de presupuestos o en 
                        un presupuesto nuevo.<br /><br />
                        También se puede seleccionar <span style="color: rgb(182, 68, 80);">NO MOVER</span> los saldos, lo cual generará los registros de cierre en el presupuesto a 
                        cerrar, pero sin crear registros de apertura en ningún otro presupuesto.<br /><br />
                    </div>
                </div>

                <div class="line"> 
                    <label class="label normal">Mover Saldos</label>
                    <select class="field full" name="preSiguiente" id="preSiguiente">
                        <option value="n">01. Crear Nuevo Presupuesto y Mover Saldos</option>
                        <option value="*" selected="selected">02. No Mover Saldos a ningún Presupuesto</option>
                        <%
                            contador = 2

                            sqlString = "SELECT Presupuesto, Nombre " & _
                                        "FROM pre_Presupuesto_Encabezado " & _
                                        "WHERE (Usuario = '" & usu & "') " & _
                                            "AND (Tipo = 'P') " & _ 
                                            "AND (Estatus = 1) " & _ 
                                            "AND (Presupuesto <> '" & pre & "') " & _
                                    "ORDER BY Presupuesto;"

                            set t = con.execute(sqlString)

                            if not (t.bof or t.eof) then
                            Do
                                contador = contador + 1
                                lblContador = right("000" & contador, 2) & ".&nbsp;Mover a&nbsp;[&nbsp;"
                                response.write "<option value='" & t("Presupuesto") & "'>" & lblContador & t("Nombre") & "&nbsp;]</option>"
                                t.MoveNext
                            Loop Until t.eof
                            end if

                            t.close: set t = nothing
                        %>
                        </select>                
                    </div>
                </div>
            </div>
        </div>

        <script type="text/javascript">
            function cerrar() {
                var npre = document.getElementById("preSiguiente").value;
                var mensaje;

                if (npre == "n") {
                    var vinculo = "pre_det_cerrar3.asp?p=<%= pre %>&np=" + npre; 
                    window.location.href = vinculo;          
                } else {
                    if (npre == "*") {
                        mensaje = "Esta seguro de Cerrar el Presupuesto '<%= pre %>' y No mover los Saldos?"
                    } else {
                        mensaje = "Esta seguro de Cerrar el Presupuesto '<%= pre %>' y mover los saldos al Presupuesto '" + npre + "'?";              
                    };

                    var confirmacion = confirm(mensaje);
                    var vinculo = "pre_det_cerrar4.asp?p=<%= pre %>&np=" + npre; 

                    if (confirmacion) {
                        window.location.href = vinculo;
                    };
                };
            }           
        </script>  
    </body>

    <% con.close: set con = nothing %>   
    <!-- #include virtual = "/core/includes/kernel/close.inc" -->
</html>