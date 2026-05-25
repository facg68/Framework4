<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>

        <title>Medios por Tipo</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  

        <%
            thisSystem = "discos"
            thisProcess = "discos.0340"
            SysLockOut

            '
            ' Init()
            '
            dim con, t, sqlString, Amo, Tipo, Usuario
            dim data, labels, Editor

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            Usuario = Request.Cookies("Usuario")
            Editor = Request.Form("cboEditor")
            Amo = Request.Form("cboAmo")
            Tipo = Request.Form("cboTipo")

            if Amo = "" then 
                CargarDefaults Usuario, "mTipo"
            else
                ActualizarDefaults Usuario, "mTipo", Tipo, Amo, Editor
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
                    Editor = "DM"
                else
                    Amo = t("SeleccionAmo")
                    Tipo = t("TipoAmo")
                    Editor = t("SeleccionMedio")
                end if

                t.close: set t = nothing
            end sub

            sub ActualizarDefaults(Usuario, Grafica, Tipo, Amo, Editor)
                if existeDefault(Usuario, Grafica) then
                    sqlString = "UPDATE discos_Graficas_Defaults " & _
                                "SET TipoAmo = " & Tipo & ", " & _
                                    " SeleccionAmo = " & Amo & ", " & _
                                    " SeleccionMedio = '" & Editor & "' " & _
                            " WHERE (Usuario = '" & Usuario & "') " & _
                                "AND (Grafica = '" & Grafica & "');"
                else
                    sqlString = "INSERT INTO discos_Graficas_Defaults(Usuario, Grafica, TipoAmo, SeleccionAmo, SeleccionMedio) " & _
                                "VALUES('" & Usuario & "', '" & Grafica & "', " & Tipo & ", " & Amo & ", '" & Editor & "');"
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

    <body plantilla="grafica" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    

        <br />

        <form id="formulario" name="formulario" method="post" action="gfx_medios_tipo.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 40%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Medios por Tipo
                </div>
                
                <div style="flex: 0 0 60%; text-align: right;">
                    <select class="no-field" name="cboEditor" id="cboEditor" onChange="Submit();">
                      <%
                        sqlString = "SELECT Codigo, Nombre " & _
                                      "FROM discos_Objetos_Clases " & _
                                  "ORDER BY Nombre;"

                        set t = con.Execute(sqlString)

                        if not (t.bof or t.eof) then
                            Do
                                response.write "<option value='" & t("Codigo") & "'"
                                    if t("Codigo") = Editor then 
                                        response.write " selected" 
                                    end if
                                response.write ">" & t("Nombre") & "</option>"
                                t.MoveNext
                            Loop Until t.eof
                        end if

                        t.close: set t = nothing
                      %>
                    </select>

                    &nbsp;&nbsp;

                    <select class="no-field" name="cboTipo" id="cboTipo" onChange="Submit()">
                        <option value="1" <% if Tipo = 1 then response.write " selected" %>>A&ntilde;o de Compra</option>
                        <option value="2" <% if Tipo = 2 then response.write " selected" %>>A&ntilde;o de Edici&oacute;n</option>
                    </select>

                    &nbsp;&nbsp;

                    <select class="no-field" name="cboAmo" id="cboAmo" onChange="Submit()">
                        <option value="1" <% if Amo = 1 then response.write " selected" %>>&Uacute;ltimos 5 A&ntilde;os</option>
                        <option value="2" <% if Amo = 2 then response.write " selected" %>>&Uacute;ltimos 10 A&ntilde;os</option>
                        <option value="3" <% if Amo = 3 then response.write " selected" %>>&Uacute;ltimos 15 A&ntilde;os</option>
                        <option value="4" <% if Amo = 4 then response.write " selected" %>>&Uacute;ltimos 20 A&ntilde;os</option>
                        <option value="5" <% if Amo = 5 then response.write " selected" %>>Todos los Registros</option>
                    </select>   
                </div>
            </div>  

            <div class="main">
                <%
                    sqlString = "discos_gfx_tipos '" & Usuario & "', " & Amo & ", " & Tipo & ", '" & Editor & "'"    
                    apexColumns "", "chart", sqlString, "Nombre", "Cantidad", "#064f8aff", 125     
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