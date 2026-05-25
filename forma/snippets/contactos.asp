<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    Snip_Width = 350    
%>

<!-- #include virtual = "/core/includes/snippets.inc" -->

<style>
    a.cont_snip_linea, a.cont_snip_linea:link, a.cont_snip_linea:visited,
    a.cont_snip_linea:focus, a.cont_snip_linea:hover, 
    a.cont_snip_linea:active { 
        font-family: Arial;
        font-size: 16px;        
        color: black; 
    }   

    a.cont_snip_linea2, a.cont_snip_linea2:link, a.cont_snip_linea2:visited,
    a.cont_snip_linea2:focus, a.cont_snip_linea2:hover, 
    a.cont_snip_linea2:active { 
        font-family: Arial;
        font-size: 16px;        
        color: rgb(76, 76, 76); 
    }

    .cont_snip_main {
        max-width: 96%;
        margin: 0.5rem auto;
        padding: 0;
        background: transparent;
        border-radius: 0;
        box-shadow: none;
        display: flex;
        font-family: sans-serif;
    }    
</style>

<%
    dim cont_snip_con, cont_snip_t, cont_snip_sqlString, cont_snip_vinculo
    dim tipo, categ, orden, dir

    cont_snip_sqlString = "SELECT DISTINCT Usuario, Codigo, Nombre, Correo, Cumple, Telefono " & _
                            "FROM ( " & _
                                    " SELECT Usuario, Codigo, Nombre, Correo, Cumple, Telefono, TipoContacto, Categ " & _
                                        " FROM con_FiltroContactos " & _
                                        " WHERE (Codigo <> '" & Request.Cookies("Usuario") & "') " & _
                                        " AND (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                        " AND (TipoContacto <> 'CU') " & _
                                        " AND (Visible = '1') " & _
                                        " AND (Telefono <> '') " & _
                                " ) as t ORDER BY Nombre ASC;"

    set cont_snip_con= Server.CreateObject("ADODB.Connection")
    cont_snip_con.open Application("Conn")
    set cont_snip_t = cont_snip_con.Execute(cont_snip_sqlString)        
%>

<div class="cont_snip_main" style="max-height: 500px;">
    <div class="tabla-wrapper">
        <table class="tabla tabla-violeta"> 
            <thead>
                <tr>
                    <th colspan="2" class="sticky" style="width:60%; padding: 5px; font-size: 12px; text-align: center;">Contactos</th>
                </tr>
            </thead>

            <tbody>
                <%
                    if not (cont_snip_t.bof or cont_snip_t.eof) then
                        Do
                            cont_snip_cuantos = cont_snip_cuantos + 1
                            vinculo = "/forma/desktop/apps/agenda/cont/cont_editar.asp?con=" & cont_snip_t("Codigo") 

                            response.write "<tr style='padding: 0;'>"
                                response.write "<td style='width:10%; text-align: center; padding: 4px;'>"
                                    response.write "<a class='cont_snip_linea' href='" & cont_snip_vinculo & "'>"%>
                                        <img src="<%= request.Cookies("usuPath") & "/fotos/" & cont_snip_t("Codigo") & "_s.jpg" %>" 
                                            onerror="this.src='/core/imagenes/misc/foto.jpg'" width="50px"><%
                                    response.write "</a>"
                                response.write "</td>"                

                                response.write "<td style='width:90%; text-align: left; padding: 0; padding-left: 10px; padding-right: 10px;'>"
                                    response.write "<a class='cont_snip_linea' href='" & vinculo & "'>"
                                        response.write "<b>" & cont_snip_t("Nombre") & "</b>"
                                    response.write "</a>"

                                    response.write "<a class='cont_snip_linea2' href='" & vinculo & "'>"
                                        response.write "<br />" & cont_snip_t("Telefono") 
                                    response.write "</a>"                                            
                                response.write "</td>"   
                            response.write "</tr>"

                            cont_snip_t.MoveNext
                        Loop Until cont_snip_t.eof
                    end if

                    cont_snip_t.close: set cont_snip_t = nothing
                %>                
            </tbody>

            <tfoot>
                <tr>
                    <td class="sticky" colspan="2" style="width: 100%; text-align: center; font-size: 12px; padding: 5px;">                
                        <%
                            Select Case cont_snip_cuantos
                                case 0: response.write "No se encontraron contactos"
                                case 1: response.write "Sólo se encontró un Contacto"
                                case else
                                    response.write "Se encontraron " & cont_snip_Cuantos &  " Contactos"                                
                            end Select
                        %>
                    </td>
                </tr>
            </tfoot>
        </table>            
    </div>
</div>
    
<% cont_snip_con.close: set cont_snip_con = nothing %>

<script>
    function contactos_init() {
        redimWindow("contactos", <%= Snip_Width %>)
    }
</script>