<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Editar Nota" %>
        <title><%= PageTitle %></title>

        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0090"
            SysLockOut

            dim con, t, p, sqlString, llave, usu, pre, vinculo
            dim Nota, multi, dia, ver, tipo, estatus, ordenado

            usu = Request.Cookies("usuario")
            pre = Request.QueryString("presupuesto")
            llave = Request.QueryString("registro")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")            

            ' Funciones y Procedimientos ---------------------------------------------------------------------------------------------------------------
                function NombrePresupuesto(Usuario, Presupuesto)
                    dim ta

                    set ta = con.Execute("SELECT nombre from pre_Presupuesto_Encabezado where (Usuario = '" & usuario & "') AND (Presupuesto = '" & presupuesto & "');")
                        if not (ta.bof or ta.eof) then
                            NombrePresupuesto = ta("nombre")
                        else
                            NombrePresupuesto = ""
                        end if
                    ta.close: set ta = nothing
                end Function

                Function FechaForm(FechaDB)
                    dim a, m, d

                    FechaForm = ""

                    if not isnull(FechaDB) then
                        d = RIGHT("00" & day(FechaDB) ,2)
                        m = RIGHT("00" & month(FechaDB), 2)
                        A = year(FechaDB)

                        FechaForm = d & "/" & m & "/" & a
                    end if
                end function

                Function HoraForm(HoraDB)
                    dim h, m

                    HoraForm = ""

                    if not isnull(HoraDB) then
                        h = LEFT(HoraDB, 2)
                        m = RIGHT(HoraDB, 2)

                        HoraForm = h & ":" & m
                    end if
                end function 
            ' ------------------------------------------------------------------------------------------------------------------------------------------
        %>             
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <div class="page-title-bar">
            <div class="title"><%= PageTitle %></div>
        </div>

        <%
            sqlString = "SELECT d.Llave, d.Fecha, d.Hora, d.Descripcion, d.Nota, d.NotaPre, d.NotaDonde " & _
                          "FROM dbo.pre_Presupuesto_Detalles AS d " & _
                         "WHERE Llave = " & llave & ";"

            set t = con.execute(sqlString)
                HoraTemp = RIGHT("0000" & t("Hora"), 4)
                Fecha = FechaForm(t("Fecha"))
                Hora = HoraForm(HoraTemp)
                Descripcion = t("Descripcion")
                Nota = t("Nota")
                NotaPre = t("NotaPre")
                NotaDonde = t("NotaDonde")
            t.close: set t = nothing                
        %>          

        <form name="formulario" id="formulario" method="post" action="pre_det_grabar_nota.asp">
            <div class="no-ver">
                <input id="Llave" name="Llave" type="text" value="<%= Llave %>" />
            </div>  

            <main>
                <div class="contenedor">
                    <br />

                    <div class="line">
                        <label>Fecha</label>
                        <input type="text" disabled value="<%= fecha & " " & hora %>">
                    </div>

                    <div class="line">
                        <label>Nota</label>
                        <% if NotaPre = 1 then %> 
                            <textarea name="txtNota" id="txtNota" rows=10 cols=80><%= Nota %></textarea>  
                        <% else %>
                            <script src="/core/lib/tinymce/tinymce.min.js"></script>
                            
                            <textarea class="editor"id="txtNota" name="txtNota"> 
                                <%= Nota %>
                            </textarea>

                            <script>
                                tinymce.init({
                                    entity_encoding : "raw",
                                    selector: '.editor',
                                    license_key: 'gpl',
                                    height: 450,
                                    branding: false,
                                    promotion: false,                                    
                                    language: 'es',
                                    language_url: '/core/includes/es.js', 
                                    plugins: 'anchor autolink charmap codesample emoticons image link lists media searchreplace table visualblocks wordcount ',
                                    toolbar: 'undo redo | blocks fontfamily fontsize | bold italic underline strikethrough | link image media table mergetags | addcomment showcomments | spellcheckdialog a11ycheck typography | align lineheight | checklist numlist bullist indent outdent | emoticons charmap | removeformat',
                                    mergetags_list: [
                                        { value: 'First.Name', title: '' },
                                        { value: 'Email', title: '' },
                                    ]
                                });
                            </script>                
                        <% end if %>
                    </div>

                    <div class="line">
                        <label>Tipo</label>
                        <select name="NotaPre" id="NotaPre">
                            <option value="0" <% if NotaPre = "0" then response.write " selected" %>>Ver Nota en formato Extendido</option>
                            <option value="1" <% if NotaPre = "1" then response.write " selected" %>>Forzar el formato de bloque en impresión</option>     
                        </select>  
                    </div>

                    <div class="line">
                        <label>Alineación</label>
                        <select name="NotaDonde" id="NotaDonde">
                            <option value="I" <% if NotaDonde = "I" then response.write " selected" %>>Imprimir Nota en Panel Izquierdo</option>
                            <option value="D" <% if NotaDonde = "D" then response.write " selected" %>>Imprimir Nota en Panel Derecho</option>     
                        </select> 
                    </div>   
                </div> 

                <br />
            </main>
        </form>

        <footer class="footer-contextual">
            <button class="footer-button" type="button" aria-label="Home" onclick="irA('/forma/mobile')">
                <i class="fa-solid fa-house"></i>
            </button>
                    
            <button class="footer-button" type="button" aria-label="Volver" onclick="volver()">
                <i class="fas fa-undo-alt"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Grabar" onclick="submit()">
                <i class="fas fa-save"></i>
            </button>            
        </footer>

        <script>
            function volver() {
                history.back();
            }

            function irA(vinculo) {
                window.location.href = vinculo;
            }            

            function submit() {
                document.getElementById("formulario").submit();
            }            
        </script>

        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>