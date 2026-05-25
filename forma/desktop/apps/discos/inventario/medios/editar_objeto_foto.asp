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
            dim con, t, sqlString, paquete, objeto, usuario, titulo

            usuario = Request.Cookies("usuario")
            paquete = Request.QueryString("p")
            objeto = Request.QueryString("o")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            sqlString = "SELECT Titulo " & _
                        "FROM discos_Objetos " & _
                        "WHERE (Usuario = '" & usuario & "') " & _
                        "AND (Paquete = '" & paquete & "') " & _
                        "AND (Objeto = '" & Objeto & "');"

            set t = con.execute(sqlString)
                titulo = t("Titulo")
            t.close: set t = nothing
        %>  

        <br />        

        <form id="frm_adjuntos" name="frm_adjuntos" action="editar_objeto_foto_upload.asp" method="post" enctype="multipart/form-data"> 
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    <%= Titulo %>
                </div>
                
                <div style="flex: 0 0 50%; text-align: right;">
                    <button class="form-btn rojo normal"  type="button" onclick="Volver('<%= paquete %>', '<%= objeto %>')">Cancelar</button>
                    <button class="form-btn verde normal" type="submit">Actualizar</button>
                </div>
            </div>    

            <br />

            <div class="main main-scroll">
                <div class="no-ver">
                    <input type="text" id="paquete" name="paquete" value="<%= paquete %>" style="width: 300px;" /> 
                    <input type="text" id="objeto" name="objeto" value="<%= objeto %>" style="width: 300px;" /> 
                </div>                        

                <div class="line">
                    <input type="file" id="File1" name="FILE1" accept=".jpg" style="width: 100%;" /> 
                </div>                

                <div class="line">
                    <% fotoPath = request.Cookies("usuPath") & "/medios/" & Objeto & ".jpg" %>

                    <img src="<%= fotoPath %>" 
                        onerror="this.src='/core/imagenes/misc/foto.jpg'" 
                        style="width:auto%; max-height: 300px;">            
                </div>
            </div>
        </form>             

        <script>
            pageReserva = 250;

            function Volver(paquete, objeto) {
                var vinculo = "editar_objeto.asp?p=" + paquete + "&o=" + objeto;
                window.location.href = vinculo;
            }        
        </script>                 
    </body>
</html>