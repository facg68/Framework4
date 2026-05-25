<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Asignar Saldos" %>
        <title><%= PageTitle %></title>

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
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <div class="page-title-bar">
            <%= PageTitle %>
        </div>

        <main>
            <br />

            <div class="contenedor">
                <div class="line">
                    Este presupuesto <u>tiene saldos pendientes</u>. Se insertarán 
                    registros de cierre en las cuenta de 
                    <span style="color: rgb(182, 68, 80);">CARTERA</span> y 
                    <span style="color: rgb(182, 68, 80);">EFECTIVO</span>, así 
                    como registros de apertura en el presupuesto que se 
                    seleccione de la lista de presupuestos o en uno nuevo.

                    <br /><br />

                    También se puede seleccionar 
                    <span style="color: rgb(182, 68, 80);">NO MOVER</span> 
                    los saldos, lo cual insertará registros de cierre, pero 
                    sin crear registros de apertura.
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
            <div>
        </main>

        <footer class="footer-contextual">
            <button class="footer-button" type="button" aria-label="Home" onclick="irA('/forma/mobile')">
                <i class="fa-solid fa-house"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Volver" onclick="volver()">
                <i class="fas fa-undo-alt"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Volver" onclick="cerrar()">
                <i class="fas fa-lock"></i>
            </button>            
        </footer>

        <script>
            function volver() {
                history.back();
            }

            function irA(vinculo) {
                window.location.href = vinculo;
            }        

            function cerrar() {
                var npre = document.getElementById("preSiguiente").value;
                var presupuesto = "<%= pre %>";

                if (npre === "n") {
                    window.location.href = "pre_det_cerrar3.asp?p=" + presupuesto + "&np=" + npre;
                    return;
                }

                var mensaje;

                if (npre === "*") {
                    mensaje = "¿Está seguro de cerrar el Presupuesto '" + presupuesto + "' y NO mover los saldos?";
                } else {
                    mensaje = "¿Está seguro de cerrar el Presupuesto '" + presupuesto + 
                            "' y mover los saldos al Presupuesto '" + npre + "'?";
                }

                Swal.fire({
                    title: "Confirmación",
                    html: "<strong>" + mensaje + "</strong>",
                    icon: "question",
                    showCancelButton: true,
                    confirmButtonText: "Cerrar",
                    cancelButtonText: "Cancelar",
                    confirmButtonColor: "#d97706",   // Naranja serio
                    cancelButtonColor: "#0d6efd",    // Azul tranquilo
                    reverseButtons: true
                }).then(function(result) {
                    if (result.isConfirmed) {
                        window.location.href = "pre_det_cerrar4.asp?p=" + presupuesto + "&np=" + npre;
                    }
                });
            }           
        </script>

        <% con.close: set con = nothing %> 
        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>