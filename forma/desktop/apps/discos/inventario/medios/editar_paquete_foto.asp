<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>

        <title>Editar Foto</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->
        <%
            thisSystem = "discos"
            thisProcess = "discos.0110"
            SysLockOut
 
            sub append(byRef Cadena, NuevaCadena)
                if NuevaCadena <> "" then
                    Cadena = Cadena & NuevaCadena
                end if
            end sub
        %>  
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->      
        <%
            dim con, t, sqlString, paquete, usuario, titulo

            usuario = Request.Cookies("usuario")
            paquete = Request.QueryString("p")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            sqlString = "SELECT Titulo " & _
                        "FROM discos_Paquetes " & _
                        "WHERE (Usuario = '" & usuario & "') " & _
                        "AND (Paquete = '" & paquete & "');"

            set t = con.execute(sqlString)
                titulo = t("Titulo")
            t.close: set t = nothing
        %>  

        <br />        

        <form id="frm_adjuntos" name="frm_adjuntos" action="editar_paquetes_foto_upload.asp" method="post" enctype="multipart/form-data">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    <%= Titulo %>
                </div>
                
                <div style="flex: 0 0 50%; text-align: right;">
                    <button class="form-btn rojo normal" type="button" onclick="Volver('<%= paquete %>')">Cancelar</button>
                    <button class="form-btn verde normal" type="submit">Actualizar</button>
                </div>
            </div>    

            <br />

            <div class="main main-scroll">
                <div class="no-ver">
                    <input type="text" id="paquete" name="paquete" value="<%= paquete %>" style="width: 300px;" /> 
                </div>   

                <div class="line">
                    <input type="file" id="File1" name="FILE1" accept=".jpg" style="width: 100%;" /> 
                </div>

                <div class="line">
                    <% fotoPath = request.Cookies("usuPath") & "/medios/" & Paquete & ".jpg" %>

                    <img src="<%= fotoPath %>" 
                        onerror="this.src='/core/imagenes/misc/foto.jpg'" 
                        style="width: auto; max-height: 300px;">
                </div>
            </div>
        </form>        

        <script>
            function Volver(paquete) {
                var vinculo = "editar.asp?m=" + paquete;
                window.location.href = vinculo;
            }        
        </script>                 
    </body>
</html>