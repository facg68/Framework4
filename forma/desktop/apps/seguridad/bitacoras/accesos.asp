<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>

        <title>Listar Bitacora de Accesos</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "seguridad"
            thisProcess = "seg.0110"
            SysLockOut
        %>    

        <style>
            select, input {
                color: black; 
                font-size: 14px; 
                background-color: rgb(226, 245, 225);  
                border-radius: 5px; 
                padding: 5px; 
                border: 1px solid rgb(200, 200, 200);                
            }

            td, th {
                padding: 5px;
            }    

            .borde {
                border: 1px solid;
                border-color: rgb(184, 184, 184);
            }   

            .filtro {
                border: 0px;
                font-size: 14px;
                color: rgb(255,255,255);
                background-color: transparent;
                text-decoration: underline;
            }
        </style>

        <%
            dim con, t, sqlString, Usuario, Sistema, Proceso, Amo, Mes
            dim sqlCommand

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            Usuario = Request.Form("cboUsuario")
            Sistema = Request.Form("cboSistema")
            Proceso = Request.Form("cboProceso")
            Amo = Request.Form("Amo")
            Mes = Request.Form("cboMes")

            if Usuario = "" then Usuario = "*"
            if Sistema = "" then Sistema = "*"
            if Proceso = "" then Proceso = "*"
            if Amo = "" then Amo = Year(Date)
            if Mes = "" then Mes = "*"

            sqlCommand = "SELECT Amo, Mes, Fecha, Sistema, Proceso, Usuario, Acceso, FechaForm, nomUsuario, nomSistema, nomProceso " & _
                          "FROM dbo.qry_seg_Bitacoras " & _
                         "WHERE (Amo = " & Amo & ") "

                if Mes <> "*"     then sqlCommand = sqlCommand & "AND (Mes = " & Mes & ") "
                if Usuario <> "*" then sqlCommand = sqlCommand & "AND (Usuario = '" & Usuario & "') "
                if Sistema <> "*" then sqlCommand = sqlCommand & "AND (Sistema = '" & Sistema & "') "
                if Proceso <> "*" then sqlCommand = sqlCommand & "AND (Proceso = '" & Proceso & "') "

            sqlCommand = sqlCommand & "ORDER BY Fecha DESC;"

            '
            ' Funciones y Procedimientos
            '      

            Sub ListarBitacora(Sistema, Proceso, Usuario, Amo, Mes)
                set tt = con.execute(sqlCommand)
                    if not (tt.bof or tt.eof) then            
                        %> 
                            <table style="width: 98%; font-size: 14px; font-family: Verdana;">
                                <tr style="background-color: rgb(85,85,85); color: rgb(255,255,255);">
                                    <td style="text-align: left; padding: 5px; width: 30%;">Usuario</td>
                                    <td style="text-align: left; padding: 5px; width: 30%;">Sistema</td>
                                    <td style="text-align: left; padding: 5px; width: 25%;">Proceso</td>
                                    <td style="text-align: center; padding: 5px; width: 15%;">Fecha</td>
                                </tr>

                                <%
                                    Do
                                        response.write "<tr>"
                                            response.write "<td style='padding: 5px; text-align:  left;'>" & tt("nomUsuario") & "</td>"
                                            response.write "<td style='padding: 5px; text-align:  left;'>" & tt("nomSistema") & "</td>"
                                            response.write "<td style='padding: 5px; text-align:  left;'>" & tt("nomProceso") & "</td>"
                                            response.write "<td style='padding: 5px; text-align: center;'>" & tt("FechaForm") & "</td>"
                                        response.write "</tr>"                                        

                                        tt.MoveNext
                                    Loop Until tt.eof

                                %>
                            </table>
                        <%
                    end if                    
                tt.close: set tt = nothing
            end Sub
        %>    
    </head>

    <body>
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    
        <div style="width: 98%; margin: auto;">
            <br />

            <form id="formulario" name="formulario" method="post" action="accesos.asp" class="cmxform form-horizontal style-form">
                <table style="width: 100%;"> 
                    <tr>
                        <td>
                            <table style="width: 100%; padding: 0px;">
                                <tr>
                                    <td style="width: 15%; text-align: left; font-size: 20px; color: rgb(6, 32, 79); font-weight: bold;">
                                        &nbsp;&nbsp;Accesos
                                    </td>

                                    <td style="width: 85%; text-align: right;">
                                        <select class="field" name="cboUsuario" id="cboUsuario" onChange="Submit()">
                                            <option value="*" <% if Usuario = "*" then response.write " selected" %>>- - Usuarios - - </option>                                           
                                            <%
                                                sqlString = "SELECT usuCodigo, usuNombre FROM seg_Usuarios WHERE usuCodigo <> 'defaults' ORDER BY usuNombre;"
                                                set t = con.execute(sqlString)
                                                    if not (t.bof or t.eof) then
                                                        Do
                                                            response.write "<option value='" & t("usuCodigo") & "' "
                                                                if Usuario = t("usuCodigo") then 
                                                                    response.write " selected" 
                                                                end if
                                                            response.write ">" & t("usuNombre") & "</option>"
                                                            t.movenext
                                                        Loop until t.eof
                                                    end if
                                                t.close: set t = nothing
                                            %>                                            
                                        </select>

                                        <select class="field" name="cboSistema" id="cboSistema" onChange="Submit()">
                                            <option value="*" <% if Sistema = "*" then response.write " selected" %>>- - Sistemas - - </option>                                           
                                            <%
                                                sqlString = "SELECT sysCodigo, sysNombre FROM seg_Sistemas ORDER BY sysNombre;"
                                                set t = con.execute(sqlString)
                                                    if not (t.bof or t.eof) then
                                                        Do
                                                            response.write "<option value='" & t("sysCodigo") & "' "
                                                                if Sistema = t("sysCodigo") then 
                                                                    response.write " selected" 
                                                                end if
                                                            response.write ">" & t("sysNombre") & "</option>"
                                                            t.movenext
                                                        Loop until t.eof
                                                    end if
                                                t.close: set t = nothing
                                            %>                                            
                                        </select>

                                        <select class="field" name="cboProceso" id="cboProceso" onChange="Submit()">
                                            <option value="*" <% if Proceso = "*" then response.write " selected" %>>- - Procesos - - </option>                                           
                                            <%
                                                if Sistema <> "*" then
                                                    sqlString = "SELECT proCodigo, proNombre FROM seg_Procesos " & _
                                                                "WHERE proSistema = '" & Sistema & "' " & _
                                                                "ORDER BY proNombre;"
                                                else
                                                    sqlString = "SELECT proCodigo, proNombre FROM seg_Procesos " & _
                                                                "WHERE proSistema = '/*' " & _
                                                                "ORDER BY proNombre;"                                                
                                                end if

                                                set t = con.execute(sqlString)
                                                    if not (t.bof or t.eof) then
                                                        Do
                                                            response.write "<option value='" & t("proCodigo") & "' "
                                                                if Proceso = t("proCodigo") then 
                                                                    response.write " selected" 
                                                                end if
                                                            response.write ">" & t("proNombre") & "</option>"
                                                            t.movenext
                                                        Loop until t.eof
                                                    end if
                                                t.close: set t = nothing
                                            %>                                            
                                        </select>                                        

                                        <input class="field tiny" style="text-align: left !important;" id="amo" name="amo" type="text" value="<%= Amo %>" onchange="Submit()" required="" style="text-align: center; padding: 6px; width: 75px;">

                                        <select class="field" name="cboMes" id="cboMes" onchange="Submit()">
                                            <option value="*" <% if Mes = "*" then response.write "required" %>>- - Todos - -</option>
                                            <option value="1" <% if Mes = "1" then response.write "required" %>>Enero</option>
                                            <option value="2" <% if Mes = "2" then response.write "required" %>>Febrero</option>
                                            <option value="3" <% if Mes = "3" then response.write "required" %>>Marzo</option>
                                            <option value="4" <% if Mes = "4" then response.write "required" %>>Abril</option>
                                            <option value="5" <% if Mes = "5" then response.write "required" %>>Mayo</option>
                                            <option value="6" <% if Mes = "6" then response.write "required" %>>Junio</option>
                                            <option value="7" <% if Mes = "7" then response.write "required" %>>Julio</option>
                                            <option value="8" <% if Mes = "8" then response.write "required" %>>Agosto</option>
                                            <option value="9" <% if Mes = "9" then response.write "required" %>>Septiembre</option>
                                            <option value="10" <% if Mes = "10" then response.write "required" %>>Octubre</option>
                                            <option value="11" <% if Mes = "11" then response.write "required" %>>Noviembre</option>
                                            <option value="12" <% if Mes = "12" then response.write "required" %>>Diciembre</option>
                                        </select>                                        

                                        &nbsp;&nbsp;                                                   
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </form>

            <br /><br />

            <% ListarBitacora Sistema, Proceso, Usuario, Amo, Mes %>
        </div>

        <% con.close: set con = nothing %>
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->        
    </body>
</html>