<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Exitos por Interpretes</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  

        <%
            thisSystem = "discos"
            thisProcess = "discos.0140"
            SysLockOut
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

    <body plantilla="tabla" reserva="185">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    
        <%  
            '
            ' Abrimos la tabla y llenamos los datos
            '
            dim con, t, sqlString, vinculo, sw
            dim InDirAu, Ver, Orden, Cuantos
            dim tt, Usuario

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            cuantos = 0
            Usuario = Request.Cookies("Usuario")

            InDirAu = Request.Form("cboInDirAu")
            Ver = Request.Form("cboVer")
            Orden = Request.Form("cboOrden")

            If Orden = "" then Orden = 4

            '
            ' Creamos la cadena de conexión, dependiendo de los
            ' datos del filtro, o generamos una cadena nueva

            sqlString = "SELECT Secuencia, Usuario, Paquete, Objeto, InDirAu, AEdicion, " & _
                                " Album, Tema, Exito, Editor " & _
                            "FROM discos_Exitos_Por_InDirAu " & _
                            "WHERE (Usuario = '" & Usuario & "') "

            if InDirAu <> "" then 
                sqlString = sqlString & "AND (InDirAu = '" & InDirAu & "') " 
            else 
                sqlString = sqlString & "AND (InDirAu = '/*/*') " 
            end if

            if Ver <> "" then   
                select case Ver
                    case 1
                        sqlString = sqlString & "AND (Exito = 1) " 
                    case 2
                        sqlString = sqlString & "AND (Exito > 0) " 
                end select
            end if

            SqlString = sqlString & "ORDER BY "

            SELECT CASE Orden
                CASE 1: sqlString = sqlString & "AEdicion ASC" 
                CASE 2: sqlString = sqlString & "Album ASC"
                CASE 3: sqlString = sqlString & "Tema ASC"
                CASE 4: sqlString = sqlString & "AEdicion DESC"
                CASE 5: sqlString = sqlString & "Album DESC"
                CASE 6: sqlString = sqlString & "Tema DESC"
            END SELECT

            set t = con.Execute(sqlString)        
        %>

        <br />

        <form id="formulario" name="formulario" method="post" action="listae.asp">
            <div style="width: 93%; margin: auto;">
                <div style="text-align: left; font-size: 25px; color: rgb(50, 50, 50); padding: 5px;">
                    Exitos por Intérprete
                </div>

                <div style="text-align: left; font-family: Ruda; font-size: 16px; color: rgb(50, 50, 50); padding: 5px;">
                    <select class="no-field" name="cboInDirAu" id="cboInDirAu" onChange="Requery();">
                        <option value="" <% if InDirAu = "" then response.write " selected" %>>&nbsp;</option>
                        <%
                            sqlString = "SELECT DISTINCT InDirAu " & _
                                        "FROM dbo.discos_Objetos " & _
                                        "WHERE (Editor = 'DM') OR (Editor = 'VM') " & _
                                        "AND (Usuario = '" & Usuario & "') " & _
                                        "ORDER BY InDirAu;"

                            set tt = con.Execute(sqlString)

                            if not (tt.bof or tt.eof) then
                                Do
                                    response.write "<option value='" & tt("InDirAu") & "'"
                                        if InDirAu = tt("InDirAu") then 
                                            response.write " selected" 
                                        end if
                                    response.write ">" & tt("InDirAu") & "</option>"
                                    tt.MoveNext
                                Loop Until tt.eof
                            end if
                        %>
                    </select>    

                    &nbsp;

                    <select class="no-field" name="cboVer" id="cboVer" onChange="Requery();">
                        <option value="1" <% if Ver = "1" then response.write " selected" %>>Exitos</option>
                        <option value="2" <% if Ver = "2" then response.write " selected" %>>Buenos</option>
                    </select>                     

                    &nbsp;

                    <select class="no-field" ame="cboOrden" id="cboOrden" onChange="Requery();">
                        <option value="1" <% if Orden = "1" then response.write " selected" %>>▲ Año</option>
                        <option value="2" <% if Orden = "2" then response.write " selected" %>>▲ Album</option>
                        <option value="3" <% if Orden = "3" then response.write " selected" %>>▲ Tema</option>
                        <option value="4" <% if Orden = "4" then response.write " selected" %>>▼ Año</option>
                        <option value="5" <% if Orden = "5" then response.write " selected" %>>▼ Album</option>
                        <option value="6" <% if Orden = "6" then response.write " selected" %>>▼ Tema</option>
                    </select> 
                </div>   
            </div>

            <div class="main">
                <% if not (t.bof or t.eof) then %>
                    <div class="line">
                        <div class="tabla-wrapper">
                            <table class="tabla tabla-green">
                                <thead>
                                    <tr style="background-color: black; color: white;"> 
                                        <th class="sticky" style="width:10%; text-align: center;">Año</th>
                                        <th class="sticky" style="width:38%; text-align: center;">Album</th>
                                        <th class="sticky" style="width:37%; text-align: center;">Tema</th>
                                        <th class="sticky" style="width:15%; text-align: center;">Exito</th>
                                    </tr>
                                </thead>

                                <tbody>
                                    <%
                                        Do
                                            cuantos = cuantos + 1
                                            vinculo = "/forma/desktop/apps/discos/inventario/medios/editar_objeto.asp?p=" & t("Paquete") & "&o=" & t("Objeto") & "&e=" & t("Editor")

                                            response.write "<tr>"
                                                response.write "<td style='text-align: center;'>"
                                                    response.write "<a class='linea' href='" & vinculo & "'>"
                                                        response.write t("aEdicion")
                                                    response.write "</a>"
                                                response.write "</td>"

                                                response.write "<td style='text-align: left; '>"
                                                    response.write "<a class='linea' href='" & vinculo & "'>"
                                                        response.write t("Album")
                                                    response.write "</a>"
                                                response.write "</td>"

                                                response.write "<td style='text-align: left; '>"
                                                    response.write "<a class='linea' href='" & vinculo & "'>"
                                                        response.write t("Tema")
                                                    response.write "</a>"
                                                response.write "</td>"                  

                                                response.write "<td style='text-align: left; '>"
                                                    response.write "<a class='linea' href='" & vinculo & "'>"
                                                        Select Case t("Exito")
                                                            case 1: response.write "Exito"
                                                            case 2: response.write "Muy Bueno"
                                                        end select
                                                    response.write "</a>"
                                                response.write "</td>" 
                                            response.write "</tr>"  

                                            t.MoveNext
                                        Loop Until t.eof

                                        t.close: set t = nothing
                                    %>
                                </tbody>

                                <tfoot>
                                    <tr>
                                        <td class="sticky" colspan="4" style="text-align: center;">
                                            <%
                                                Select Case cuantos
                                                    case 0: response.write "No se encontraron registros"
                                                    case 1: response.write "Sólo se encontró un registro"
                                                    case else
                                                        response.write "Se encontraron " & cuantos & " registros"
                                                end Select
                                            %>                            
                                        </td>                                
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                    </div>
                <% end if %>
            </div>
        </form>

        <br />

        <script type="text/javascript">
            function Requery() {
                document.getElementById("formulario").submit();
            } 
        </script>       
     
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
        <% con.close: set con = nothing  %>
    </body>
</html>