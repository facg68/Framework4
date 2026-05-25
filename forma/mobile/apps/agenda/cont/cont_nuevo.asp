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
        %>           
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <%
            dim con, t, sqlString, cbox, c, nomContacto, nuevo
            dim ver, tipo, categ, orden1, orden2

            dim usuario, codigo, tipoContacto, primerNombre, segundoNombre, primerApellido, segundoApellido, correoElectronico, fechaCumple
            dim empresa, telefonoEmpresa, sitioWeb, pais, provincia, ciudad, direccion, notas, deSistema, visible, arbol, signo
            
            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            usuario = Request.Cookies("usuario")
            codigo = "nuevo"
            tipoContacto = "PE"
            DeSistema = 0
            Visible = 1
            Arbol = 99
            Signo = 99

            nomContacto = ""   
        %>          

        <div class="page-title-bar">
            <%= "Nuevo Contacto" %>
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
                        <label>Categoría:</label>
                        <select name="tipoContacto" id="tipoContacto">
                            <option value="PE" <% if tipoContacto = "PE" then response.write " selected" %>>Persona / Contacto Principal</option>
                            <option value="ES" <% if tipoContacto = "ES" then response.write " selected" %>>Establecimiento / Locales</option>                            
                            <option value="CU" <% if tipoContacto = "CU" then response.write " selected" %>>Cuenta / Cuenta</option>
                        </select> 
                    </div>

                    <div class="line">
                        <label>Primer Nombre</label>
                        <input id="primerNombre" name="primerNombre" type="text" value="" />
                    </div>

                    <div class="line">
                        <label>Segundo Nombre</label>
                        <input id="segundoNombre" name="segundoNombre" type="text" value="" />
                    </div>

                    <div class="line">
                        <label>Primer Apellido</label>
                        <input id="primerApellido" name="primerApellido" type="text" value="" />
                    </div>

                    <div class="line">
                        <label>Segundo Apellido</label>
                        <input id="segundoApellido" name="segundoApellido" type="text" value="" />
                    </div>                                                

                    <div class="line">
                        <label>Correo</label>
                        <input id="correoElectronico" name="correoElectronico" type="text" placeholder="direccion@server.com" />
                    </div>            

                    <div class="line">
                        <label>Cumpleaños</label>
                        <input id="fechaCumple" name="fechaCumple" type="text" placeholder="dd/mm" />
                    </div> 

                    <div class="line">
                        <label>Teléfono Principal</label>
                        <input id="telefono" name="telefono" type="text" value="" />
                    </div>  

                    <div class="line">
                        <label>Labora En</label>
                        <input id="empresa" name="empresa" type="text" />
                    </div> 

                    <div class="line">
                        <label>Teléfono Empresa</label>
                        <input id="telefonoEmpresa" name="telefonoEmpresa" type="text" />
                    </div> 

                    <div class="line">
                        <label>Sitio Web</label>
                        <input id="sitioWeb" name="sitioWeb" type="text" />
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
                                        response.write "<option value='" & cbox("Codigo") & "'>" & cbox("Nombre") & "</option>"
                                        cbox.MoveNext
                                    Loop Until cbox.eof
                                end if

                                cbox.close: set cbox = nothing
                            %>
                        </select>  
                    </div> 

                    <div class="line">
                        <label>Provincia</label>
                        <input id="provincia" name="provincia" type="text" />
                    </div> 

                    <div class="line">
                        <label>Ciudad</label>
                        <input id="ciudad" name="ciudad" type="text"  />
                    </div>

                    <div class="line">
                        <label>Dirección</label>
                        <textarea name="txtAreaDireccion" id="txtAreaDireccion" rows=5 cols=60></textarea> 
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

            <button class="footer-button" type="button" aria-label="Volver" onclick="grabar()">
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
                nTel = nTel.replace("+","*");

                document.getElementById("tel").value = nTel;
                document.getElementById("form_transaccion").submit();
            }    

            mask(document.getElementById('fechaCumple'), ['99/99']);               
        </script>

        <% con.close: set con = nothing %>    
        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>