<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Gastos Mensuales Por A&ntilde;o</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  

        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0130"
            SysLockOut

            dim con, t, tt, sqlString, data, labels, amo, Usuario

            Usuario = Request.Cookies("Usuario")
            amo = request.Form("amo")
            if amo = "" then amo = YEAR(NOW())   

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")     

            sub CalcularDatosGfx(Amo)
                sqlString = "exec pre_gfx_GastosMensuales '" & Request.Cookies("Usuario") & "', " & Amo & ";"
                set tt = con.execute(sqlString)

                if not (tt.bof or tt.eof) then
                    data = ""
                    labels = ""
                    
                    Do
                        if data <> "" then
                            data = data & ", " & tt("Mensual")
                            labels = labels & ", '" & tt("Nombre") & "'"
                        else
                            data = tt("Mensual")
                            labels = "'" & tt("Nombre") & "'"
                        end if                

                        tt.MoveNext
                    Loop until tt.eof
                end if

                tt.close: set tt = nothing  
            end sub 
        %>    
    </head>

    <body plantilla="grafica" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->   

        <br />

        <form id="formulario" name="formulario" method="post" action="mensuales.asp">
            <div style="display: flex; justify-content: space-between; width: 94%; margin: auto;">
                <div style="flex: 0 0 75%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Grafica de Gastos Mensuales (por Año)
                </div>
                
                <div style="flex: 0 0 25%; text-align: right;">
                    Año&nbsp;
                    <input class="no-field" style="width: 75px; text-align: center;" id="Amo" name="Amo" type="text" value="<%= Amo %>" onChange="Submit()" required />
                </div>
            </div>        

            <div class="main">
                <%
                    sqlString = "exec pre_gfx_GastosMensuales '" & Usuario & "', " & Amo 
                    apexColumns "", "chart", sqlString, "Nombre", "Mensual", "#064f8aff", 125     
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