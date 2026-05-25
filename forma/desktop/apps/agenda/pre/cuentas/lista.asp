<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->   
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0100"
            SysLockOut    

            sub ActualizarTodasLasCuentas()
                dim c 

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    c.Execute("exec pre_Enc_ActualizarSaldoCuentas '" & Request.Cookies("Usuario") & "'")
                c.close: set c = nothing
            end sub

            function CuantasCuentas(TipoCuenta, GrupoCuenta)
                dim c, tt

                set c = server.CreateObject("ADODB.Connection")
                    c.open Application("Conn")

                    set tt = c.Execute("SELECT COUNT(*) AS Cuantas " & _
                                        "FROM pre_Cuentas " & _
                                        "WHERE Usuario = '" & Request.Cookies("Usuario") & "'" & _
                                        "AND (TipoCuenta = '" & TipoCuenta & "') " & _
                                        "AND (Grupo = '" & GrupoCuenta & "');")

                    if (tt.bof or tt.eof) then
                        CuantasCuentas = 0
                    else
                        CuantasCuentas = tt("Cuantas")
                    end if
                    
                    tt.close: set tt = nothing
                c.close: set c = nothing   
            end function

            function EnUso(Cuenta)
                dim cc, tt

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                    set tt = cc.execute("SELECT COUNT(*) AS Cuantos " & _
                                            "FROM pre_Presupuesto_Detalles AS d " & _
                                            "WHERE (Usuario = '" & Request.Cookies("Usuario") & "' AND CuentaOrigen = '" & Cuenta & "') " & _
                                            "OR (Usuario = '" & Request.Cookies("Usuario") & "' AND CuentaDestino = '" & Cuenta & "');")
                        if tt("Cuantos") > 0 then
                            EnUso = 1
                        else
                            EnUso = 0
                        end if
                    tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function      
        %>        

        <style>
            a.linea:link,
            a.linea:visited,
            a.linea:focus,
            a.linea:hover,
            a.linea:active {
                color: black !important;
            }
            
            td { 
                padding: 2 !important;
                vertical-align: middle !important;
            }
        </style>                 
    </head>

    <body plantilla="tabla" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            ActualizarTodasLasCuentas
            
            '
            ' Abrimos la tabla y llenamos los datos
            '

            dim con, t, sqlString, vinculo, sw
            dim Grupo, orden, cuantos

            cuantos = 0

            Grupo = Request.Form("cboVerGrupo")
            if Grupo = "" then
                Grupo = request.QueryString("g")
                Orden = request.QueryString("o")
            end if

            if Grupo = "" then Grupo = "A"
            if Orden = "" then Orden = 1
            
            '
            ' Creamos la cadena de conexión, dependiendo de los
            ' datos del filtro, o generamos una cadena nueva

            sqlString = "SELECT * FROM pre_Cuentas_ListaCuentasUsuario " & _
                        "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " 

            if Grupo <> "*" then sqlString = sqlString & "AND (Grupo = '" & Grupo & "') " 

            SqlString = sqlString & "ORDER BY "

            SELECT CASE Orden
                CASE 1: sqlString = sqlString & "Codigo " & oDir
                CASE 2: sqlString = sqlString & "Seccion " & oDir
                CASE 3: sqlString = sqlString & "Nombre " & oDir
                CASE 4: sqlString = sqlString & "Tipo " & oDir
                CASE 5: sqlString = sqlString & "Monto " & oDir
                CASE 6: sqlString = sqlString & "Anualidad " & oDir
                CASE 7: sqlString = sqlString & "NombreContacto " & oDir
                CASE 8: sqlString = sqlString & "Clase " & oDir
            END SELECT

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")
            set t = con.Execute(sqlString)        
        %>      

        <br />

        <form id="formulario" name="formulario" method="post" action="lista.asp">
            <div style="display: flex; justify-content: space-between; width: 93%; margin: auto;">
                <table style="width: 100%;">
                    <tr>
                        <td style="width: 15%; font-size: 24px;">
                            Cuentas
                        </td>

                        <td style="width: 85%; text-align: right;">
                            <select class="field" name="cboVerGrupo" id="cboVerGrupo" onChange="Requery();">
                                <option value="A" <% if Grupo = "A" then response.write " selected" %>>Cuentas Activas</option>
                                <option value="W" <% if Grupo = "W" then response.write " selected" %>>Cuentas En Espera</option>
                                <option value="S" <% if Grupo = "S" then response.write " selected" %>>Cuentas Archivadas</option>
                                <option value="*" <% if Grupo = "*" then response.write " selected" %>>Todas las Cuentas</option>                    
                            </select>                                

                            &nbsp;&nbsp;

                            <a href="cuentas_editar.asp?c=">
                                <button type="button" class="form-btn verde" style="width: 150px; font-size: 16px; color:white;">Nueva Cuenta</button>
                            </a>

                            &nbsp;&nbsp;   

                            <a onclick="NoAplicadas()">
                                <button type="button" class="form-btn azul" style="width: auto; font-size: 16px; color: white; padding: 10px;">Borrar Transacciones No Aplicadas</button>
                            </a>        

                            &nbsp;&nbsp;

                            <a onclick="CerrarTransacciones()">
                                <button type="button" class="form-btn rojo" style="width: auto; font-size: 16px; color: white; padding: 10px;">Cierre de Transacciones</button>  
                            </a>                            

                            &nbsp;&nbsp;
                        </td>                        
                    </tr>
                </table>
            </div> 
        </form>      

        <div class="main">
            <div class="tabla-wrapper">
                <table class="tabla tabla-green">
                    <thead>
                        <tr>
                            <th class="sticky" style="width:11%; text-align: center; padding: 10px;" onClick="ordenar('1');">Codigo</th>
                            <th class="sticky" style="width: 8%; text-align: center; padding: 10px;" onClick="ordenar('2');">Seccion</th>
                            <th class="sticky" style="width:23%; text-align: center; padding: 10px;" onClick="ordenar('3');">Nombre</th>
                            <th class="sticky" style="width: 8%; text-align: center; padding: 10px;" onClick="ordenar('4');">Tipo</th>
                            <th class="sticky" style="width:10%; text-align: center; padding: 10px;" onClick="ordenar('5');">Monto</th>
                            <th class="sticky" style="width: 5%; text-align: center; padding: 10px;" onClick="ordenar('6');">Anual</th>
                            <th class="sticky" style="width:20%; text-align: center; padding: 10px;" onClick="ordenar('7');">Contacto</th>
                            <th class="sticky" style="width:10%; text-align: center; padding: 10px;" onClick="ordenar('8');">Clase</th>
                            <th class="sticky" style="width: 5%; text-align: center;">&nbsp;</th>                            
                        </tr>
                    </thead>

                    <tbody>
                        <%
                            if not (t.bof or t.eof) then
                                response.write "<tbody>"

                                Do
                                    cuantos = cuantos + 1
                                    vinculo = "cuentas_editar.asp?c=" & t("Codigo")

                                    if Len( Trim(t("Estatus")) ) = 2 then
                                        subClase = "azul"
                                    else
                                        if left(t("Estatus"), 1) = "C" then    
                                            subClase = "verde"
                                        else
                                            subClase = "rojo"
                                        end if                                    
                                    end if

                                    response.write "<tr"
                                        if subClase <> "" then 
                                            response.write " class='tr-" & subClase & "'"
                                        end if
                                    response.write ">"
                                        response.write "<td style='text-align: left;'>"
                                            response.write "<a class='linea' href='" & vinculo & "'>"
                                                response.write t("Codigo")
                                            response.write "</a>"
                                        response.write "</td>"            

                                        response.write "<td style='text-align: left;'>"
                                            response.write "<a class='linea' href='" & vinculo & "'>"
                                                response.write t("Seccion")
                                            response.write "</a>"
                                        response.write "</td>"    

                                        response.write "<td style='text-align: left;'>"
                                            response.write "<a class='linea' href='" & vinculo & "'>"
                                                response.write t("Nombre")
                                            response.write "</a>"
                                        response.write "</td>"    

                                        response.write "<td style='text-align: left;'>"
                                            response.write "<a class='linea' href='" & vinculo & "'>"
                                                response.write t("Tipo")
                                            response.write "</a>"
                                        response.write "</td>"    

                                        response.write "<td style='text-align: right;'>"
                                            response.write "<a class='linea' href='" & vinculo & "'>"
                                                response.write FormatNumber(t("Monto"))
                                            response.write "</a>"
                                        response.write "</td>"   

                                        response.write "<td style='text-align: center;'>"
                                            response.write "<a class='linea' href='" & vinculo & "'>"
                                                response.write t("Anualidad")
                                            response.write "</a>"
                                        response.write "</td>"  

                                        response.write "<td style='text-align: left;'>"
                                            response.write "<a class='linea' href='" & vinculo & "'>"
                                                response.write t("NombreContacto")
                                            response.write "</a>"
                                        response.write "</td>" 

                                        response.write "<td style='text-align: left;'>"
                                            response.write "<a class='linea' href='" & vinculo & "'>"
                                                response.write t("Clase")
                                            response.write "</a>"
                                        response.write "</td>" 

                                        response.write "<td style='text-align: center;'>"
                                            if EnUso(t("Codigo")) = 1 then
                                                ctaEstatus = "disabled"
                                            end if
                                                
                                            %>
                                                <a class="linea" onclick="borrar('<%= t("Codigo") %>')" <%= ctaEstatus %>>
                                                    <button type="button" class="form-btn rojo">
                                                        <i class="fa fa-trash fa_xl" title='Borrar Cuenta'></i>
                                                    </button>
                                                </a>
                                            <%
                                        response.write "</td>" 
                                    response.write "</tr>"

                                    t.MoveNext
                                Loop Until t.eof

                                response.write "</tbody>"
                            end if

                            t.close: set t = nothing
                        %>                    
                    </tbody>

                    <tfoot>
                        <tr>
                            <td class="sticky" colspan="9" style="text-align: center;">
                                <%
                                    Select Case cuantos
                                        case 0: response.write "No hay Cuentas"
                                        case 1: response.write "Una Cuenta"
                                        case else
                                            response.write "Se encontraron " & Cuantos & " Cuentas"
                                    end Select
                                %>                            
                            </td>
                        </tr>
                    </tfoot>
                </table>
            </div>      
        </div>
  
        <br />

        <script type="text/javascript">
            function Requery() {
                document.getElementById("formulario").submit();
            }

            function ordenar(campo) {
                var grupo = document.getElementById("cboVerGrupo").value;
                var vinculo = "lista.asp?g=" + grupo + "&o=" + campo;
        
                window.location.href = vinculo;
            }

            function borrar(cuenta) {
                var confirmacion = confirm("Está seguro de borrar la cuenta " + cuenta + "?");
                var vinculo = "cuentas_borrar.asp?c=" + cuenta;

                if (confirmacion) {
                    window.location.href = vinculo;
                } else {
                    alert("Proceso Cancelado.");        
                }        
            }

            function NoAplicadas() {
                var confirmacion = confirm("¿Está seguro de borrar las transacciones no aplicadas?");
                var direccion = "cuentas_borrar_no_aplicadas.asp";

                if (confirmacion) {
                    window.location.href = direccion;
                }          
            }

            function CerrarTransacciones() {
                var confirmacion = confirm("¿Está seguro de cerrar las transacciones?");
                var direccion = "cuentas_cerrar.asp";

                if (confirmacion) {
                    window.location.href = direccion;
                }          
            } 
        </script>   

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
        <% con.close: set con = nothing %> 
    </body>
</html>