<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>

        <title>Editar Foto</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0050"
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
            dim con, t, sqlString, cbox, c, nomContacto, nuevo
            dim ver, tipo, categ, orden1, orden2

            dim usuario, codigo
            
            usuario = Request.Cookies("usuario")
            codigo = Request.QueryString("con")

            ver = Request.QueryString("v")
            tipo = Request.QueryString("t")
            categ = Request.QueryString("c")
            orden1 = Request.QueryString("o1")
            orden2 = Request.QueryString("o2")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            sqlString = "SELECT codigo, primerNombre, segundoNombre, primerApellido, segundoApellido " & _
                            "FROM con_Contactos as c " & _
                            "WHERE (Usuario = '" & usuario & "') " & _
                            "AND (Codigo = '" & codigo & "');"

            set t = con.execute(sqlString)

                primerNombre = t("primerNombre")
                segundoNombre = t("segundoNombre")
                primerApellido = t("primerApellido")        
                segundoApellido = t("segundoApellido")

                nomContacto = ""
                if PrimerNombre    <> "" then append nomContacto, " " & PrimerNombre        
                if SegundoNombre   <> "" then append nomContacto, " " & SegundoNombre
                if PrimerApellido  <> "" then append nomContacto, " " & PrimerApellido
                if SegundoApellido <> "" then append nomContacto, " " & SegundoApellido

            t.close: set t = nothing
        %>  

        <br />        

        <form id="frm_adjuntos" name="frm_adjuntos" action="cont_upload_foto.asp" method="post" enctype="multipart/form-data">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    <%= nomContacto %>
                </div>
                
                <div style="flex: 0 0 50%; text-align: right;">
                    <button class="form-btn rojo large" type="button" onclick="Volver()">Cancelar</button>
                    <button class="form-btn verde large" type="submit">Actualizar Foto</button>
                </div>
            </div>    

            <br />

            <div class="main main-scroll">
                <div class="no-ver">
                    <input type="text" id="contacto"    name="contacto" value="<%= Codigo %>"   style="width: 300px;" /> 
                    <input type="text" id="ver"         name="ver"      value="<%= ver %>"      style="width: 300px;" /> 
                    <input type="text" id="tipo"        name="tipo"     value="<%= tipo %>"     style="width: 300px;" /> 
                    <input type="text" id="categ"       name="categ"    value="<%= categ %>"    style="width: 300px;" /> 
                    <input type="text" id="orden1"      name="orden1"   value="<%= orden1 %>"   style="width: 300px;" /> 
                    <input type="text" id="orden2"      name="orden2"   value="<%= orden2 %>"   style="width: 300px;" /> 

                </div>   

                <div class="line">
                    <input type="file" id="File1" name="FILE1" accept=".jpg" style="width: 100%;" /> 
                </div>

                <div class="line">
                    <label class="label xxl">
                        <img src="<%= request.Cookies("usuPath") & "/fotos/" & Codigo & ".jpg" %>" 
                            onerror="this.src='/core/imagenes/misc/foto.jpg'" 
                            style="width:100%; height:auto;">                                              
                    </label>
                </div>
            </div>
        </form>        

        <script>
            function Volver() {
                var vinculo = "con_editar.asp?con=<%= codigo %>&v=<%= ver %>&t=<%= tipo %>&c=<%= categ %>&o1=<%= orden1 %>&o2=<%= orden2 %>";
                window.location.href = vinculo;
            }      
        </script>                 
    </body>
</html>