<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>

        <title>Gastos Anuales</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "discos"
            thisProcess = "discos.0305"
            SysLockOut


            '
            ' Init()
            '

            dim con, t, sqlString, Amo, Tipo, Usuario
            dim data, labels

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            Usuario = Request.Cookies("Usuario")
            Amo = Request.Form("cboAmo")
            Tipo = Request.Form("cboTipo")

            if Amo = "" then 
                CargarDefaults Usuario, "gAnuales"
            else
                ActualizarDefaults Usuario, "gAnuales", Tipo, Amo
            end if


            '
            ' Funciones y Procedimientos
            '

            sub CargarDefaults(Usuario, Grafica)
                sqlString = "SELECT TipoAmo, SeleccionAmo, SeleccionMedio " & _
                            "FROM discos_Graficas_Defaults " & _
                            "WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Grafica = '" & Grafica & "');"

                set t = con.execute(sqlString)

                if (t.bof or t.eof) then
                    Amo = 1
                    Tipo = 1
                else
                    Amo = t("SeleccionAmo")
                    Tipo = t("TipoAmo")
                end if

                t.close: set t = nothing
            end sub

            sub ActualizarDefaults(Usuario, Grafica, Tipo, Amo)
                if existeDefault(Usuario, Grafica) then
                sqlString = "UPDATE discos_Graficas_Defaults " & _
                            "SET TipoAmo = " & Tipo & ", " & _
                                " SeleccionAmo = " & Amo & _
                        " WHERE (Usuario = '" & Usuario & "') " & _
                            "AND (Grafica = '" & Grafica & "');"
                else
                sqlString = "INSERT INTO discos_Graficas_Defaults(Usuario, Grafica, TipoAmo, SeleccionAmo, SeleccionMedio) " & _
                            "VALUES('" & Usuario & "', '" & Grafica & "', " & Tipo & ", " & Amo & ", NULL);"
                end if

                con.execute(sqlString)
            end sub

            function existeDefault(Usuario, Grafica)
                sqlString = "SELECT * FROM discos_Graficas_Defaults WHERE Usuario = '" & Usuario & "' AND Grafica = '" & Grafica & "';"

                set t = con.execute(sqlString)

                if (t.bof or t.eof) then
                existeDefault = false
                else
                existeDefault = true
                end if

                t.close: set t = nothing
            end function
        %>   
    </head>

    <body plantilla="panoramica" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->   

        <br />

        <form id="formulario" name="formulario" method="post" action="gfx_gastos_anuales.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 40%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Gastos por Año
                </div>
                
                <div style="flex: 0 0 60%; text-align: right;">
                    <select class="no-field" name="cboTipo" id="cboTipo" onChange="Submit()">
                        <option value="1" <% if Tipo = 1 then response.write " selected" %>>Año de Compra</option>
                        <option value="2" <% if Tipo = 2 then response.write " selected" %>>Año de Edición</option>
                    </select>

                    &nbsp;&nbsp;                  

                    <select class="no-field" name="cboAmo" id="cboAmo" onChange="Submit()">
                        <option value="1" <% if Amo = 1 then response.write " selected" %>>Ultimos 5 Años</option>
                        <option value="2" <% if Amo = 2 then response.write " selected" %>>Ultimos 10 Años</option>
                        <option value="3" <% if Amo = 3 then response.write " selected" %>>Ultimos 15 Años</option>
                        <option value="4" <% if Amo = 4 then response.write " selected" %>>Ultimos 20 Años</option>
                        <option value="5" <% if Amo = 5 then response.write " selected" %>>Todos los Registros</option>
                    </select>   
                </div>
            </div>        

            <div class="main">
                <%
                    sqlString = "exec discos_gfx_gastos '" & Usuario & "', " & Tipo & ", " & Amo
                    apexPanoramic "", sqlString, "Año", "Total", "#008FFB", 500
                %>                         
            </div>
        </form>
  
        <br /><br />   

        <script>
            function Submit() {
                document.getElementById("formulario").submit(); 
            } 
        </script>

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
        <% con.close: set con = nothing %> 
    </body>
</html>