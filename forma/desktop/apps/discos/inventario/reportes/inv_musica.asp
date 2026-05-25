<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Reporte de Musica por Interpretes </title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  

        <%
            thisSystem = "discos"
            thisProcess = "discos.0135"
            SysLockOut

            dim cc, tt, sqlString, Usuario, sqlTabla, sqlElementos
            dim cuantas, contador, items(4) 

            Usuario = Request.Cookies("usuario")

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")

            function CuantosItems(ConsultaSQL)
                dim ctable, sqlCommand

                sqlCommand = "SELECT COUNT(*) AS Cuantos FROM (" & ConsultaSQL & ") AS t;"

                set ctable = cc.execute(sqlCommand)
                    if not (ctable.bof or ctable.eof) then
                        CuantosItems = ctable("Cuantos")
                    end if
                ctable.close: set ctable = nothing
            end function

            Sub DistribuirEnSecciones(elementos)
                Dim base, division, total, i

                total = elementos

                if total >= 4 then
                    division = cdbl(total / 4.00)

                    if division - int(total / 4) = 0 then 
                        base = Int(total / 4)
                    else
                        base = Int(total / 4) + 1
                    end if

                    For i = 1 To 4
                        if total > base then
                            items(i) = base
                            total = total - base
                        else
                            items(i) = total
                        end if
                    Next
                else
                    For i = 1 To 4
                        if total > 0 then
                            items(i) = 1
                            total = total - 1
                        else
                            items(i) = 0
                        end if
                    Next                    
                end if
            End Sub  

            Sub DibujarTablaElementos(ConsultaSQL, Prefijo)
                set tt = cc.execute(ConsultaSQL & " ORDER BY Nombre;")
                
                if not (tt.bof or tt.eof) then
                    cuantas = CuantosItems(ConsultaSQL)

                    if cuantas > 0 then
                        DistribuirEnSecciones(cuantas)

                        response.write "<table style='border: none;'>"
                            response.write "<tr style='border: none;'>"
                                for i = 1 To 4
                                    response.write "<td style='vertical-align: top; width: 25%;'>"
                                        if items(i) > 0 then
                                            for e = 1 to items(i)
                                                %>
                                                    <p>
                                                        <input type='checkbox' id="<%= Prefijo & tt("Codigo") %>" name="<%= Prefijo & tt("Codigo") %>" value="1" checked />
                                                        <label>
                                                            <%
                                                                if cuantas > 0 then
                                                                    if tt("Nombre") = "-" then
                                                                        response.write "No Definida"
                                                                    else
                                                                        response.write tt("Nombre")
                                                                    end if

                                                                    cuantas = cuantas - 1
                                                                end if
                                                            %>
                                                        </label>
                                                    </p>
                                                <%

                                                tt.movenext
                                            next 
                                        else
                                            response.write "&nbsp;"
                                        end if
                                    response.write "</td>"
                                next
                            response.write "<tr>"
                        response.write "</table>"
                    end if
                end if            
            End Sub
        %>     

		<style>
            p {
                margin-bottom: 10px;
                line-height: 1.5;
            }  
		</style>         
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    
        <br />

        <form name="form_transaccion" id="form_transaccion" method="post" action="inv_musica_rep.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 65%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Reporte de Música por Intérpretes
                </div>
                
                <div style="flex: 0 0 35%; text-align: right;">
                    <button class="form-btn verde normal" type="button" onclick="informe()">Informe</button>                
                </div>
            </div>    

            <div class="main main-scroll">
                <div class="line">
                    <label class="label normal">Agrupación</label>
                    <label class="label full section">
                        <table style='border: none;'>
                            <tr style='border: none;'>
                                <td style="width: 25%;">
                                    <p>
                                        <input type='radio' id='chk_ruptura' name='chk_ruptura' value='1' checked='checked'>
                                        <label>Inicial</label>
                                    </p>
                                </td>

                                <td style="width: 25%;">
                                    <p>
                                        <input type='radio' id='chk_ruptura' name='chk_ruptura' value='2'>
                                        <label>Nombre</label>
                                    </p>
                                </td>

                                <td style="width: 25%;">&nbsp;</td>
                                <td style="width: 25%;">&nbsp;</td>
                            </tr>
                        </table>
                    </label>
                </div>

                <div class="line">
                    <label class="label normal">Colección</label>
                    <label class="label full section">
                        <%
                            sqlTabla =  "SELECT DISTINCT Carpeta AS Codigo, NombreCarpeta AS Nombre " & _
                                        "FROM dbo.discos_Rep_Musica_InDirAu " & _
                                        "WHERE (Usuario = '" & Usuario & "') " 

                            DibujarTablaElementos sqlTabla, "c"
                        %>                                                 
                    </label>
                </div>  

                <div class="line">
                    <label class="label normal">Forma</label>
                    <label class="label full section">
                        <%
                            sqlTabla =  "SELECT DISTINCT Forma AS Codigo, NombreForma AS Nombre " & _
                                        "FROM dbo.discos_Rep_Musica_InDirAu " & _
                                        "WHERE (Usuario = '" & Usuario & "') " 

                            DibujarTablaElementos sqlTabla, "f"
                        %>                                            
                    </label>
                </div>

                <div class="line">
                    <label class="label normal">Tienda</label>
                    <label class="label full section">
                        <%
                            sqlTabla = "SELECT DISTINCT Tienda AS Codigo, NombreTienda AS Nombre " & _
                                       "FROM dbo.discos_Rep_Musica_InDirAu " & _
                                       "WHERE (Usuario = '" & Usuario & "') "

                            DibujarTablaElementos sqlTabla, "t"
                        %>                      
                    </label>
                </div>  
            </div>
        </form>

        <br /><br />                                 

        <script>
            function informe() {
                document.getElementById("form_transaccion").submit(); 
            }
        </script>

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->    
        <% cc.close: set cc = nothing %>
    </body>
</html>