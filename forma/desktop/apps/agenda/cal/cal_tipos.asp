<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <title>Editar Tipos de Calendarios</title>    
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->             
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0030"
            SysLockOut
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

            .col2 { flex: 1; }

            td {
                position: relative;
            }

            td input[type="checkbox"] {
                position: relative;
                top: 1px; /* ajuste fino, casi zen */
            }        

            .chk-line input[type="checkbox"] {
                vertical-align: 2.5px;
            }    
        </style>           
    </head>

    <body plantilla="tabla" reserva="225">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            dim con, t, sqlString, cuantos
            dim Codigo, Nombre, Vacio, usu

            usu = Request.Cookies("Usuario")
            cuantos = 0

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            sqlString = "SELECT Usuario, Codigo, Nombre, PorDefecto, DeSistema, ColorFont, Seleccionado " & _
                        "FROM cal_Calendarios " & _
                        "WHERE (Usuario = '" & usu & "') " & _
                    "ORDER BY Nombre ASC;"

            set t = con.execute(sqlString)
        %>          

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                Tipos de Calendario
            </div>
            
            <div style="flex: 0 0 50%; text-align: right;">
                <button class='form-btn verde large' 
                        type='button' 
                        onclick="submit_Estados()"  >
                        Actualizar Estados
                </button>
            </div>
        </div>        

        
        <div class="main" style="width: 95%;">
            <form name="frm_Estado" id="frm_Estado" method="post" action="cal_tipos_actualizar.asp">
                <div class="line">
                    <div class="tabla-wrapper">
                        <table class="tabla tabla-green">
                            <thead>
                                <tr>
                                    <th class="sticky" style="width: 95%;">Nombre</th>
                                    <th class="sticky" style="width:  5%;">&nbsp</th>
                                </tr>
                            </thead>

                            <tbody>                    
                                <%
                                    if not (t.bof or t.eof) then
                                        Cuantos = 0

                                        Do
                                            Cuantos = Cuantos + 1
                                            estado = ""                                                
                                            seleccionado = ""

                                            if t("DeSistema") = 1 then estado = " disabled"
                                            if t("Seleccionado") = 1 then seleccionado = " checked"


                                            %>
                                                <tr>
                                                    <td>
                                                        <label class="chk-line">
                                                            <input type="checkbox" name="ck_<%= t("Codigo") %>" value="1" <%= seleccionado %>>
                                                            &nbsp;
                                                            <span><%= t("Nombre") %></span>
                                                        </label>
                                                    </td>

                                                    <td>
                                                        <button type="button" class="form-btn rojo" onclick="borrar('<%= t("Codigo") %>')" <%= estado %>>
                                                            <i class="fa fa-trash fa-xl" title="Borrar Tipo de Calendario"></i>
                                                        </button>
                                                    </td>
                                                </tr>
                                            <%

                                            t.MoveNext
                                        Loop Until t.eof
                                    end if
                                %>
                            </tbody>

                            <tfoot>
                                <tr>
                                    <td class="sticky" style="text-align: center;" colspan="2">
                                        <%
                                            response.write "Se encontraron " & cuantos & " Calendarios"
                                        %>
                                    </td>
                                </tr>
                            </tfoot>
                        </table>
                    </div>
                </div>
            </form>


            <form name="form_nuevo" id="form_nuevo" method="post" action="cal_nuevo_tipo.asp">
                <div class="fila">
                    <div class="col col1">Nuevo Calendario</div>

                    <div class="col col2">
                        <input class="field" style="width: 100%;" type="text" id="nuevoNombre" name="nuevoNombre" >
                    </div>

                    <div class="col col3">
                        <button class="form-btn verde " type="button" onclick="submit_Nuevo()">
                            <i class="fa fa-save fa-xl" title="Añadir"></i>
                        </button>   
                    </div>
                </div>              
            </form>
        </div>                
  
        <br /><br />   

        <script type="text/javascript">
            function submit_Nuevo() {
                document.getElementById("form_nuevo").submit();    
            }

            function submit_Estados() {
                document.getElementById("frm_Estado").submit();    
            }

            function borrar(codigo) {
                var confirmacion = confirm("Está seguro de borrar el Tipo " + codigo + "?");
                var vinculo = "cal_borrar_calendario.asp?c=" + codigo;

                if (confirmacion) {
                    window.location.href = vinculo;
                } else {
                    alert("Proceso Cancelado.");
                }        
            }                
        </script>

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>