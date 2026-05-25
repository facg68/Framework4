<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Subir Objeto</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->
        <%
            thisSystem = "anuncios"
            thisProcess = "anu.0200"
            SysLockOut

            usuario = Request.Cookies("usuario")

            response.cookies("edit_Panfleto") = Request.QueryString("cu")
            response.cookies("edit_estatusPanfleto") = Request.QueryString("e")
            response.cookies("edit_ordenadoPor") = Request.QueryString("op")
        %>

        <style>
            table, tr, td, th, tbody {
                width: 100%;
            }

            td, th {
                padding: 5px;
                font-size: 16px;
            }   

            .top {
                background-color: rgb(71,71,71);
                color: white;
            }             
        </style>   

        <%
            dim con, t, sqlString, Panfleto, usuario, NombreAttachment
            dim estatusPanfleto, ordenadoPor, uploadsDirVar

            Response.Expires = -1
            Server.ScriptTimeout = 600
            
            Session.CodePage  = 65001        

            sub append(byRef Cadena, NuevaCadena)
                if NuevaCadena <> "" then
                    Cadena = Cadena & NuevaCadena
                end if
            end sub

            function TituloPanfleto()    
                dim cc, tt, ssql

                ssql = "SELECT Nombre FROM seg_Panfletos WHERE CU = '" & Request.Cookies("edit_Panfleto") & "';"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.Open Application("Conn")
                    set tt = cc.execute(ssql)
                        TituloPanfleto = tt("Nombre")
                    tt.close: set tt = nothing
                cc.close: set cc = nothing                 
            end function

            function TituloObjeto()    
                dim cc, tt, ssql

                ssql = "SELECT Objeto FROM seg_Panfletos WHERE CU = '" & Request.Cookies("edit_Panfleto") & "';"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.Open Application("Conn")
                    set tt = cc.execute(ssql)
                        TituloObjeto = tt("Objeto")
                    tt.close: set tt = nothing
                cc.close: set cc = nothing                 
            end function            

            function NombrePanfleto()
                dim c, t, sqlStr

                sqlStr = "SELECT Nombre FROM seg_Panfletos WHERE (CU = '" & Request.Cookies("edit_Panfleto") & "');"

                set c = Server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    set t = c.execute(sqlStr)
                        NombrePanfleto = trim(t("Nombre"))
                    t.close: set t = nothing
                c.close: set c = nothing      
            end function            

            function OutputForm()
                %>
                    <br />

                    <form id="frm_adjuntos" name="frm_adjuntos" 
                          style="width:98%; margin: auto;"
                          action="subir_objeto2.asp" method="post" enctype="multipart/form-data"> 

                        <table style="width: 95%; margin: auto;">
                            <tr>
                                <td style="width: 55%; text-align: left; font-size: 18px;">
                                    <span style="font-size: 24px;">
                                        <%= NombrePanfleto() %><br />
                                        <span style="font-size: 18px;">
                                            PANFLETO PDF <%= TituloObjeto() %>
                                        </span>
                                    </span>
                                </td>

                                <td style="width: 45%; text-align: right;">
                                    <button class="form-btn rojo normal" onclick="Volver()" type="button">Cancelar</button>                                                
                                    <button class="form-btn verde normal" type="submit">Actualizar</button>
                                </td>
                            </tr>
                        </table>

                        <div class="main main-scroll">
                            <div class="line">
                                <table>
                                    <tr>
                                        <td style="text-align:center; width:25%;">
                                            <% fotoPath = "/forma/desktop/apps/anuncios/pdf/" & NombrePanfleto() %>

                                            <img src="<%= fotoPath %>" 
                                                onerror="this.src='/forma/desktop/apps/anuncios/imagenes/foto.jpg'" 
                                                style="width:100%; height:auto;">
                                        </td>

                                        <td style="width:5%;">&nbsp;</td>

                                        <td style="width:65%;">
                                            <input class="field xxl" type="file" id="attach1" name="attach1" accept=".pdf,application/pdf" /> 
                                            <input class="no-ver" type="text" id="Panfleto" name="Panfleto" value="<%= Request.Cookies("edit_Panfleto") %>" /> 
                                        </td>

                                        <td style="width:5%;">&nbsp;</td>
                                    </tr>
                                </table>
                            </div>
                        </div>
                    </form>             
                <%
            end function
        %>

        <script>
            function onSubmitForm() {
                var formDOMObj = document.frmSend;

                if (formDOMObj.attach1.value == "")
                    alert("Please press the Browse button and pick a file.")
                else
                    return true;
                return false;
            }
        </script>
    </head>

    <body style="background-color: rgb(235, 235, 235);">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <% OutputForm() %>

        <script>
            pageReserva = 175;

            function Volver() {
                var vinculo = "lista.asp" + "?e=<%= Request.Cookies("edit_estatusPanfleto") %>&op=<%= Request.Cookies("edit_ordenadoPor") %>";
                window.location.href = vinculo;          
            }           
        </script>        
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</HTML>