<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Actualizar Foto" %>
        <title><%= PageTitle %></title>

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

        <style>
            .contenedor img {
                max-height: 400px;     /* límite absoluto */
                max-width: 90%;        /* límite relativo al contenedor */
                
                height: auto;
                width: auto;
                object-fit: contain;                
                border-radius: 18px;
                border: none;

                box-shadow: 0 6px 6px rgba(0, 0, 0, 0.60);
                transition: transform 0.25s ease, box-shadow 0.25s ease;
            }
        </style>         
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <%
            dim con, t, sqlString, cbox, c, nomContacto, nuevo
            dim usuario, codigo
            
            usuario = Request.Cookies("usuario")
            codigo = Request.QueryString("con")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            function NombreContacto(CodUsuario, CodContacto)
                dim tt, ssql

                NomContacto = ""
                ssql = "SELECT codigo, primerNombre, primerApellido " & _
                       "FROM con_Contactos " & _
                       "WHERE (Usuario = '" & CodUsuario & "') " & _
                       "AND (Codigo = '" & CodContacto & "');"

                set tt = con.execute(ssql)
                    if not (tt.bof or tt.eof) then
                        if tt("primerNombre")   <> "" then append NomContacto, " " & tt("primerNombre")        
                        if tt("primerApellido") <> "" then append NomContacto, " " & tt("primerApellido")
                    end if
                tt.close: set tt = nothing   

                NombreContacto = NomContacto         
            end function

            sub append(byRef Cadena, NuevaCadena)
                if NuevaCadena <> "" then
                    Cadena = Cadena & NuevaCadena
                end if
            end sub            
        %>  

        <div class="page-title-bar">
            <%= NombreContacto(Usuario, Codigo) %>
        </div>

        <form id="formulario" name="formulario" action="cont_upload_foto.asp" method="post" enctype="multipart/form-data">
            <div class="no-ver">
                <input type="text" id="contacto"    name="contacto" value="<%= Codigo %>"   style="width: 300px;" /> 
            </div>   

            <main>
                <br />

                <div class="contenedor" style="text-align: center;">
                    <% fotoObjeto = request.Cookies("usuPath") & "/fotos/" & Codigo & ".jpg" %>

                    <img class="img-limitada" 
                        src="<%= fotoObjeto %>" 
                        onerror="this.src='/core/imagenes/misc/foto.jpg'" >
                <div>

                <br />

                <div class="line">
                    <input type="file" id="File1" name="FILE1" accept=".jpg" style="width: 100%;" /> 
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
                const form = document.getElementById("formulario");
                const input = document.getElementById("File1");

                if (!input.files || input.files.length === 0) {
                    Swal.fire({
                        icon: 'warning',
                        title: 'Archivo requerido',
                        text: 'Debe seleccionar una imagen JPG.',
                        confirmButtonText: 'Entendido',
                        background: '#f2f2f2',
                        customClass: {
                            popup: 'swal-f4-popup',
                            title: 'swal-f4-title',
                            htmlContainer: 'swal-f4-text'
                        }
                    });

                    return;
                }

                form.requestSubmit();
            }   
        </script>

        <% con.close: set con = nothing %>
        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>