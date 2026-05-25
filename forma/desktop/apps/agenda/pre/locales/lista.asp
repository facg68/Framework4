<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <title>Editar Locales Monetarios</title>    
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->             
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0120"
            SysLockOut


            function LocalActivo(LocalMonetario)
                dim cc, tt, sqlString

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                    set tt = cc.execute("SELECT dbo.pre_Cuentas_MonedaInstancias('" & LocalMonetario & "') AS Cuantos;")
                        if tt("Cuantos") > 0 then
                            LocalActivo = 1
                        else
                            LocalActivo = 0
                        end if
                    tt.close: set tt = nothing
                cc.close: set cc = nothing        
            end function
        %>

        <style>
            .fila {
                display: flex;
                align-items: center;
                gap: 10px;
            }

            .col1 {
                white-space: nowrap;
                font-weight: bold;
                min-width: fit-content;
            }

            .col2 { flex: 0 0 10%; }
            .col3 { flex: 1; }
            .col4 { flex: 0 0 5%; }
            .col5 {flex: 0 0 10%; }
            .col6 { flex: 0 0 60px; }

            a.linea, a.linea:link, a.linea:visited,
            a.linea:focus, a.linea:hover, 
            a.linea:active { color: black; }                
        </style>           
    </head>

    <body plantilla="tabla" reserva="225">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            dim con, t, sqlString, cbox, cuantos
            dim Local, Nombre, Simbolo, Formula
        
            cuantos = 0

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            sqlString = "select Local, Nombre, Simbolo, Formula " & _
                        "from seg_Cripto_NumParse_Locales " & _
                        "where local <> 'NUM' " & _
                        "order by Nombre;"

            set t = con.execute(sqlString)
        %>           

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                Locales Monetarios
            </div>
            
            <div style="flex: 0 0 50%; text-align: right;">
                <button type="button" class="form-btn naranja" style="width: 250px;" onclick="cargar()">Cargar Datos desde Internet</button>
            </div>
        </div>        

        <div class="main" style="width: 95%;">
            <div class="line">
                <div class="tabla-wrapper">
                    <table class="tabla tabla-green">
                        <thead>
                            <tr>
                                <th class="sticky" style="width: 15%; text-align: center;">Local</td>
                                <th class="sticky" style="width: 45%; text-align: center;">Nombre</td>
                                <th class="sticky" style="width: 15%; text-align: center;">Simbolo</td>
                                <th class="sticky" style="width: 15%; text-align: center;">Formula</td>
                                <th class="sticky" style="width: 10%">&nbsp</td>
                            </tr>
                        </thead>

                        <tbody>                    
                            <%
                                if not (t.bof or t.eof) then
                                    Do
                                        cuantos = cuantos + 1
                                        vinculo = "locales_editar.asp?l=" & t("Local")

                                        response.write "<tr>"
                                            response.write "<td style='text-align: center;'>"
                                                response.write "<a class='linea' href='" & vinculo & "'>"
                                                response.write t("Local")
                                                response.write "</a>"
                                            response.write "</td>"                              

                                            response.write "<td>"
                                                response.write "<a class='linea' href='" & vinculo & "'>"
                                                response.write t("Nombre")
                                                response.write "</a>"
                                            response.write "</td>"                              

                                            response.write "<td style='text-align: center;'>"
                                                response.write "<a class='linea' href='" & vinculo & "'>"
                                                response.write t("Simbolo")
                                                response.write "</a>"
                                            response.write "</td>"                              

                                            response.write "<td style='text-align: center;'>"
                                                response.write "<a class='linea' href='" & vinculo & "'>"
                                                response.write FormatNumber(t("Formula"))
                                                response.write "</a>"
                                            response.write "</td>"                              

                                            response.write "<td style='text-align: center;'>"
                                                if LocalActivo(t("Local")) = 0 then
                                                    estatus = ""
                                                else
                                                    estatus = " disabled"
                                                end if %>

                                                <a onclick="borrar('<%= t("Local") %>', '<%= t("Nombre") %>')" <%= estatus %>>
                                                        <button class="form-btn rojo<%= estatus %>">
                                                            <i class="fa fa-trash fa-xl" title='Borrar localidad'></i>
                                                        </button>
                                                </a><%                              
                                            response.write "</td>"
                                        response.write "</tr>"                        

                                        t.MoveNext
                                    Loop Until t.eof
                                end if
                            %>
                        </tbody>

                        <tfoot>
                            <tr>
                                <td class="sticky" style="text-align: center;" colspan="5">
                                    <%
                                        if cuantos = 0 then
                                            response.write "No se ha encontrado locales Monetarios."
                                        else
                                            response.write "Se han encontrado " & cuantos & " locales Monetarios."
                                        end if
                                    %>
                                </td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>

            <form name="form_transaccion" id="form_transaccion" method="post" action="locales_nuevo.asp">
                <div class="fila">
                    <div class="col1">Nuevo Local Monetario</div>

                    <div class="col2">
                        <input class="field" style="width: 100%; text-align: center;" id="local" name="local" type="text" value="" placeholder="Nuevo" />
                    </div>

                    <div class="col3">
                        <input class="field" style="width: 100%; text-align: left;" id="Nombre" name="Nombre" type="text" value="" placeholder="Nuevo Local Monetario" />
                    </div>

                    <div class="col4">
                        <input class="field" style="width: 100%; text-align: center;" id="simbolo" name="simbolo" type="text" placeholder="AAA"/>
                    </div>

                    <div class="col5">
                        <input class="field" style="width: 100%; text-align: right;" id="formula" name="formula" type="number" placeholder="0.00"/>
                    </div>

                    <div class="col6">
                        <button class="form-btn verde " type="submit">
                            <i class="fa fa-save fa-xl" title="Añadir"></i>
                        </button>   
                    </div>
                </div>             
            </form>
        </div>
  
        <br /><br />   

        <script type="text/javascript">
            function borrar(codigo, nombre) {
                var confirmacion = confirm("Está seguro de borrar el local " + nombre + "?");
                var vinculo = "locales_borrar.asp?lm=" + codigo;

                if (confirmacion) {
                    window.location.href = vinculo;
                } else {
                    alert("Proceso Cancelado.");        
                }        
            }     

            function cargar() {
                var confirmacion = confirm("Cargar valores desde Internet sobreescribe todos los valores en la base de datos. ¿Está seguro de realizar el proceso?");
                var vinculo = "locales_cargar.asp"

                if (confirmacion) {
                    window.location.href = vinculo;
                } else {
                    alert("Proceso Cancelado.");        
                }        
            }   

            mask(document.getElementById('local'), ['_________']);           
            mask(document.getElementById('simbolo'), ['AAA']);           
        </script>

        <% con.close: set con = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>