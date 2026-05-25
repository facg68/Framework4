<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">

        <title>Subir Objeto</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->
        <%
            thisSystem = "discos"
            thisProcess = "discos.0205"
            SysLockOut
        %>  
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->       
        <%
            dim con, t, sqlString, cbox
            dim usuario, forma, imagen, NombreForma
            
            usuario = Request.Cookies("usuario")
            forma = Request.QueryString("c")


            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            sqlString = "SELECT Usuario, Forma, Nombre, Icono_Forma " & _
                        "FROM dbo.discos_Formas " & _
                        "WHERE (Usuario = '" & Usuario & "') " & _
                        "AND (Forma = '" & Forma & "');"

            set t = con.execute(sqlString)
                NombreForma = t("Nombre")
                imagen = t("Icono_Forma")
            t.close: set t = nothing
        %>  

        <br />        

        <form id="frm_adjuntos" name="frm_adjuntos" 
              style="width:98%; margin: auto;"
              action="forma_upload_foto.asp" method="post" 
              enctype="multipart/form-data"> 

            <table style="width: 95%; margin: auto;">
                <tr>
                    <td style="width: 55%; text-align: left; font-size: 18px;">
                        <span style="font-size: 24px;">
                            <%= NombreForma %>
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
                                <% fotoPath = lcase(Request.Cookies("usuPath")) & "/discos/" & imagen %>
                                <img src="<%= fotoPath %>" 
                                    onerror="this.src='/core/imagenes/misc/foto.jpg'" 
                                    style="width:100%; height:auto;">
                            </td>

                            <td style="width:5%;">&nbsp;</td>

                            <td style="width:65%;">
                                <input class="field xxl" type="file" id="File1" name="File1" accept=".gif" /> 
                                <input class="no-ver" type="text" id="Forma" name="Forma" value="<%= Forma %>" />                                                          
                            </td>

                            <td style="width:5%;">&nbsp;</td>
                        </tr>
                    </table>
                </div> 
            </div>
        </form> 

        <script>
            function Volver() {
                var vinculo = "lista.asp";
                window.location.href = vinculo;
            }        
        </script>                 
    </body>
</html>