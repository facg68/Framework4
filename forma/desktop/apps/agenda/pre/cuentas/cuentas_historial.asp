<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->    

        <%  
            function LocalMonetarioUsuario()
                dim fcon, f

                set fcon = Server.CreateObject("ADODB.Connection")
                fcon.Open Application("Conn")
                    set f = fcon.execute("SELECT isnull(usuLocal, 'US') AS usuLocal FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';")
                        LocalMonetarioUsuario = f("usuLocal")
                    f.close: set f = nothing
                fcon.close: set fcon = nothing
            end function

            function preLocalDestino(presupuesto)
                dim fcon, f

                set fcon = Server.CreateObject("ADODB.Connection")
                fcon.Open Application("Conn")
                    set f = fcon.execute("SELECT MonedaDestino FROM pre_Presupuesto_Encabezado WHERE Presupuesto = '" & presupuesto & "' AND Usuario = '" & Request.Cookies("Usuario") & "'")
                        preLocalDestino = f("MonedaDestino")
                    f.close: set f = nothing
                fcon.close: set fcon = nothing
            end function

            function HoraLista(Fechatabla)
                dim t, d, m, a, puntero, k

                a = Year(Fechatabla)
                m = RIGHT("00" & Month(Fechatabla), 2)
                d = RIGHT("00" & Day(Fechatabla), 2)

                puntero = -1
                for k = 1 to len(Fechatabla)
                    if mid(Fechatabla, k, 1) = " " then
                        if puntero < 0 then
                            puntero = k
                        end if
                    end if
                next

                on error resume next

                t = mid(Fechatabla, puntero + 1, 12)

                if err.number <> 0 then
                    response.write FechaTabla & "<br/>"
                end if
                
                HoraLista = d & "/" & m & "/" & a & "<br/>" & t
            end function   

            function NombreCuenta(Codigo) 
                dim cc, tt, sqlString

                sqlString = "SELECT Nombre FROM pre_Cuentas " & _
                            "WHERE Usuario = '" & Request.Cookies("Usuario") & "' " & _
                            "AND Codigo = '" & Codigo & "';"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                    set tt = cc.execute(sqlString)
                        if not (tt.bof or tt.eof) then
                            NombreCuenta = tt("Nombre")
                        else
                            NombreCuenta = ""
                        end if
                    tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function        
        %>                
    </head>

    <body plantilla="tabla" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            dim con, t, sqlString, c

            usu = Request.Cookies("usuario")
            c = Request.QueryString("c")
            sqlString = "pa_pre_Cuentas_Historial_Cerrado '" & usu & "','" & c  & "', '" & LocalMonetarioUsuario() & "'"

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")
            set t = con.execute(sqlString)
        %>          

        <br />

        <div style="display: flex; justify-content: space-between; width: 92%; margin: auto;">
            <div style="flex: 0 0 85%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                Transacciones Historicas de la Cuenta <%= NombreCuenta(c) %>&nbsp(<%= c %>)
            </div>
            
            <div style="flex: 0 0 15%; text-align: right;">
                <button type="button" class="form-btn normal verde" onClick="vinculo('cuentas_editar.asp?c=<%= c %>')">Volver</button>
            </div>
        </div>        

        <div class="main" style="width: 95%;">
            <div class="line">
                <div class="tabla-wrapper">
                    <table class="tabla tabla-carbon">
                        <thead>
                            <tr>
                                <th class="sticky" style="text-align: center; width: 12%;">Fecha</th>
                                <th class="sticky" style="text-align: center; width: 32%;">Presupuesto</th>
                                <th class="sticky" style="text-align: center; width: 32%;">Descripcion</th>
                                <th class="sticky" style="text-align: center; width: 14%;">Contacto</th>
                                <th class="sticky" style="text-align: center; width: 10%;">Monto</th>      
                            </tr>
                        </thead>

                        <tbody>
                            <% 
                                if not (t.bof or t.eof) then
                                    cuantos = 0

                                    do 
                                        cuantos = cuantos + 1

                                        response.write "<tr>"
                                            response.write "<td style='text-align: center;'>" & HoraLista(t("FechaHora")) & "</td>"
                                            response.write "<td>" & t("Nombre") & "</td>"
                                            response.write "<td>" & t("Descripcion") & "</td>"
                                            response.write "<td>" & t("Contacto") & "</td>"
                                            response.write "<td style='text-align: right;'>" & FormatNumber(t("Valor")) & "</td>"
                                        response.write "</tr>"

                                        t.MoveNext
                                    Loop until t.eof
                                end if
                            %>                                                            
                        </tbody>

                        <tfoot>
                            <tr>
                                <td colspan="5" style="text-align: center;" class="sticky">
                                    <%
                                        if cuantos > 0 then
                                            if cuantos = 1 then
                                                response.write "Sólo se encontró una transacción histórica"
                                            else
                                                response.write "Se encontraron " & cuantos & " transacciones históricas"
                                            end if
                                        else
                                            response.write "No se encontraron transacciones históricas"
                                        end if
                                    %>                                
                                </td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
        </div>
  
        <br /><br />   

        <script type="text/javascript">
            function vinculo(direccion) {
                window.location.href = direccion;
            }            
        </script>

        <% con.close: set con = nothing %>
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>