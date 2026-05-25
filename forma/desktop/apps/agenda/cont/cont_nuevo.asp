<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->   
        <title>Editar Contactos</title>
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0050"
            SysLockOut
        %>     
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            dim con, t, sqlString, cbox, c, nomContacto, nuevo
            dim ver, tipo, categ, orden1, orden2

            dim usuario, codigo, tipoContacto, primerNombre, segundoNombre, primerApellido, segundoApellido, correoElectronico, fechaCumple
            dim empresa, telefonoEmpresa, sitioWeb, pais, provincia, ciudad, direccion, notas, deSistema, visible, arbol, signo
            
            ver = Request.QueryString("v")
            tipo = Request.QueryString("t")
            categ = Request.QueryString("c")
            orden1 = Request.QueryString("o1")
            orden2 = Request.QueryString("o2")

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
        <br />

        <form name="form_transaccion" id="form_transaccion" method="post" action="cont_grabar.asp" style="width:100%;">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 70%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Nuevo Contacto
                </div>
                
                <div style="flex: 0 0 30%; text-align: right;">
                    <button class='form-btn rojo normal'  type='button' onclick="Volver()">Cancelar</button>&nbsp;&nbsp;
                    <button class='form-btn verde normal' type='button' onclick="Enviar()">Grabar</button>&nbsp;&nbsp;
                </div>
            </div>        

            <div class="main main-scroll">
                <div class="no-ver">
                    <input id="cod" name="cod" type="text" value="<%= codigo %>" />
                    <input id="usu" name="usu" type="text" value="<%= usuario %>" />
                    <input id="tel" name="tel" type="text" value="" />            
                </div>

                <div class="line">
                    <label class="label normal">Categoría</label>

                    <select class="field large" name="tipoContacto" id="tipoContacto">
                        <option value="PE" <% if tipoContacto = "PE" then response.write " selected" %>>Persona / Contacto Principal</option>
                        <option value="ES" <% if tipoContacto = "ES" then response.write " selected" %>>Establecimiento / Locales</option>                            
                        <option value="CU" <% if tipoContacto = "CU" then response.write " selected" %>>Cuenta / Cuenta</option>
                    </select> 
                </div>

                <div class="line">
                    <label class="label normal">Primer Nombre</label>
                    <input class="field large" id="primerNombre" name="primerNombre" type="text" value="" />
                </div>

                <div class="line">
                    <label class="label normal">Segundo Nombre</label>
                    <input class="field large" id="segundoNombre" name="segundoNombre" type="text" value="" />
                </div>

                <div class="line">
                    <label class="label normal">Primer Apellido</label>
                    <input class="field large" id="primerApellido" name="primerApellido" type="text" value="" />
                </div>

                <div class="line">
                    <label class="label normal">Segundo Apellido</label>
                    <input class="field large" id="segundoApellido" name="segundoApellido" type="text" value="" />
                </div>                                                

                <div class="line">
                    <label class="label normal">Correo</label>
                    <input class="field normal" id="correoElectronico" name="correoElectronico" type="text" placeholder="direccion@server.com" />
                </div>            

                <div class="line">
                    <label class="label normal">Cumpleaños</label>
                    <input class="field tiny" id="fechaCumple" name="fechaCumple" type="text" placeholder="dd/mm" />
                </div> 

                <div class="line">
                    <label class="label normal">Teléfono Principal</label>
                    <input class="field small" id="telefono" name="telefono" type="text" value="" />
                </div>  

                <div class="line">
                    <label class="label normal">Labora En</label>
                    <input class="field xl" id="empresa" name="empresa" type="text" />
                </div> 

                <div class="line">
                    <label class="label normal">Teléfono Empresa</label>
                    <input class="field small" id="telefonoEmpresa" name="telefonoEmpresa" type="text" />
                </div> 

                <div class="line">
                    <label class="label normal">Sitio Web</label>
                    <input class="field xl" id="sitioWeb" name="sitioWeb" type="text" />
                </div> 

                <div class="line">
                    <label class="label normal">País</label>
                    <select class="field normal" name="cboPais" id="cboPais" >
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
                    <label class="label normal">Provincia</label>
                    <input class="field large" id="provincia" name="provincia" type="text" />
                </div> 

                <div class="line">
                    <label class="label normal">Ciudad</label>
                    <input class="field large" id="ciudad" name="ciudad" type="text"  />
                </div>

                <div class="line">
                    <label class="label normal">Dirección</label>
                    <textarea class="field xl" name="txtAreaDireccion" id="txtAreaDireccion" rows=5 cols=60></textarea> 
                </div>                
            </div>
        </form>
  
        <br /><br />   

        <script type="text/javascript">
            function Volver() {
                window.location.href = "lista.asp";
            }    

            function Enviar() {
                var nTel = document.getElementById("telefono").value;
                nTel = nTel.replace("+","*");
                document.getElementById("tel").value = nTel;
                document.getElementById("form_transaccion").submit();
            }    

            mask(document.getElementById('fechaCumple'), ['99/99']);    
        </script>

        <% con.close: set con = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>