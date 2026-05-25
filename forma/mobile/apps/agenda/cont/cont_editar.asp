<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Editar Contactos" %>
        <title><%= PageTitle %></title>        

        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0050"
            SysLockOut

            dim con, t, sqlString, cbox, c, nomContacto, nuevo
            dim ver, tipo, categ, orden1

            dim usuario, codigo, primerNombre, segundoNombre, primerApellido, segundoApellido, correoElectronico, fechaCumple
            dim telefonoPrincipal, empresa, telefonoEmpresa, sitioWeb, pais, provincia, ciudad, direccion
            
            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")


            function PrimerTelefono(contacto)
                dim c, t, sqlString, usu, valor 
                usu = Request.Cookies("Usuario")

                if contacto <> "" then
                    sqlString = "SELECT TOP (1) Telefono " & _
                                "FROM dbo.con_Contactos_Telefonos " & _
                                "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                "AND (Codigo = '" & contacto & "') " & _
                                "ORDER BY Tipo;"

                    set c = Server.CreateObject("ADODB.Connection")
                    c.open Application("Conn")
                        set t = c.execute(sqlString)
                            if (t.bof or t.eof) then
                                PrimerTelefono = ""
                            else
                                PrimerTelefono = t("Telefono")
                            end if
                        t.close: set t = nothing
                    c.close: set c = nothing            
                end if
            end function              
        %>           
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <%
            usuario = Request.Cookies("usuario")
            codigo = Request.QueryString("con")

            if codigo = "" then response.redirect "lista.asp"

            sqlString = "SELECT codigo, primerNombre, segundoNombre, primerApellido, segundoApellido, " & _
                            " correoElectronico, fechaCumple, empresa, telefonoEmpresa, sitioWeb, pais, provincia, " & _
                            " ciudad, direccion " & _
                        "FROM con_Contactos as c " & _
                        "WHERE (Usuario = '" & usuario & "') " & _
                        "AND (Codigo = '" & codigo & "');"  

            set t = con.execute(sqlString)
                if (t.bof or t.eof) then 
                    response.redirect "lista.asp"
                else
                    primerNombre = t("primerNombre")
                    segundoNombre = t("segundoNombre")
                    primerApellido = t("primerApellido")
                    segundoApellido = t("segundoApellido")
                    correoElectronico = t("correoElectronico")
                    fechaCumple = t("fechaCumple")
                    telefonoPrincipal = PrimerTelefono(codigo)
                    empresa = t("empresa")
                    telefonoEmpresa = t("telefonoEmpresa")
                    sitioWeb = t("sitioWeb")
                    pais = t("pais")
                    provincia = t("provincia")
                    ciudad = t("ciudad")
                    direccion = t("direccion")
                end if
            t.close: set t = nothing
        %>          

        <div class="page-title-bar">
            <%= PrimerNombre & " " & PrimerApellido %>
        </div>

        <form name="form_transaccion" id="form_transaccion" method="post" action="cont_grabar.asp">
            <div class="no-ver">
                <input id="cod" name="cod" type="text" value="<%= codigo %>" />
                <input id="usu" name="usu" type="text" value="<%= usuario %>" />
                <input id="tel" name="tel" type="text" value="" />   
            </div>

            <main>          
                <br />

                <div class="contenedor">
                    <div class="line">
                        <label>Primer Nombre</label>
                        <input id="primerNombre" name="primerNombre" type="text" value="<%= primerNombre %>" />
                    </div>

                    <div class="line">
                        <label>Segundo Nombre</label>
                        <input id="segundoNombre" name="segundoNombre" type="text" value="<%= segundoNombre %>" />
                    </div>

                    <div class="line">
                        <label>Primer Apellido</label>
                        <input id="primerApellido" name="primerApellido" type="text" value="<%= primerApellido %>" />
                    </div>

                    <div class="line">
                        <label>Segundo Apellido</label>
                        <input id="segundoApellido" name="segundoApellido" type="text" value="<%= segundoApellido %>" />
                    </div>                                                

                    <div class="line">
                        <label>Correo</label>
                        <input id="correoElectronico" name="correoElectronico" type="text" value="<%= correoElectronico %>" placeholder="direccion@server.com" />
                    </div>            

                    <div class="line">
                        <label>Cumpleaños</label>
                        <input id="fechaCumple" name="fechaCumple" type="text" value="<%= fechaCumple %>" placeholder="dd/mm" />
                    </div> 

                    <div class="line">
                        <label>Teléfono Principal</label>
                        <input id="telefono" name="telefono" type="text" value="<%= telefonoPrincipal %>" />
                    </div>  

                    <div class="line">
                        <label>Labora En</label>
                        <input id="empresa" name="empresa" type="text" value="<%= empresa %>" />
                    </div> 

                    <div class="line">
                        <label>Teléfono Empresa</label>
                        <input id="telefonoEmpresa" name="telefonoEmpresa" type="text" value="<%= telefonoEmpresa %>" />
                    </div> 

                    <div class="line">
                        <label>Sitio Web</label>
                        <input id="sitioWeb" name="sitioWeb" type="text" value="<%= sitioWeb %>" />
                    </div>  

                    <div class="line">
                        <label>País</label>
                        <select name="cboPais" id="cboPais" >
                            <option value="*">- - Todos - -</option>
                            <%
                                sqlString ="SELECT Codigo, Nombre FROM seg_Paises ORDER BY Nombre ASC;"

                                set cbox = con.execute(sqlString)

                                if not (cbox.bof or cbox.eof) then
                                    Do
                                        response.write "<option value='" & cbox("Codigo") & "'"
                                            if pais = cbox("Codigo") then
                                                response.write " selected"
                                            end if
                                        response.write ">" & cbox("Nombre") & "</option>"
                                        
                                        cbox.MoveNext
                                    Loop Until cbox.eof
                                end if

                                cbox.close: set cbox = nothing
                            %>
                        </select>  
                    </div> 

                    <div class="line">
                        <label>Provincia</label>
                        <input id="provincia" name="provincia" type="text" value="<%= provincia %>" />
                    </div> 

                    <div class="line">
                        <label>Ciudad</label>
                        <input id="ciudad" name="ciudad" type="text"  value="<%= ciudad %>" />
                    </div>

                    <div class="line">
                        <label>Dirección</label>
                        <textarea name="txtAreaDireccion" id="txtAreaDireccion" rows=5 cols=60><%= direccion %></textarea> 
                    </div>  
                </div>
            </main>
        </form>

        <footer class="footer-contextual">
            <button class="footer-button" type="button" aria-label="Home" onclick="irA('/forma/mobile')">
                <i class="fa-solid fa-house"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Volver" onclick="volver()">
                <i class="fas fa-undo-alt"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Editar Foto" onclick="foto('<%= codigo %>')">
                <i class="fa-solid fa-camera"></i>
            </button>            

            <button class="footer-button" type="button" aria-label="Grabar Cambios" onclick="grabar()">
                <i class="fa-solid fa-floppy-disk"></i>
            </button>            
        </footer>

        <script>
            function volver() {
                history.back();
            }
            
            function irA(vinculo) {
                window.location.href = vinculo;
            }

            function grabar() {
                var nTel = document.getElementById("telefono").value;
                nTel = nTel.replace("*","+");

                document.getElementById("tel").value = nTel;
                document.getElementById("form_transaccion").submit();
            }    

            function foto(CodEmpleado) {
                Swal.fire({
                    title: "¿Desea cambiar la foto de este contacto?",
                    text: "Se reemplazará la foto actual",
                    icon: "question",
                    showCancelButton: true,
                    confirmButtonColor: "#3085d6",
                    cancelButtonColor: "#d33",
                    confirmButtonText: "Subir Fotografía"
                }).then((result) => {
                    if (result.isConfirmed) {
                            var vinculo ="cont_foto.asp?con=" + CodEmpleado;
                            window.location.href = vinculo;
                    }
                });               
            }

            mask(document.getElementById('fechaCumple'), ['99/99']);               
        </script>

        <% con.close: set con = nothing %>    
        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>