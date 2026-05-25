<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Filtrar Contactos" %>
        <title><%= PageTitle %></title>

        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0050"
            SysLockOut


            dim cont_con, cont_t, cont_sqlString, cont_vinculo

            set cont_con = Server.CreateObject("ADODB.Connection")
            cont_con.open Application("Conn")        

            visibilidad = Request.Cookies("cont_v")
            tipo = Request.Cookies("cont_t")
            categ = Request.Cookies("cont_c")            

            if tipo = "" then tipo = "PE"
            if categ = "" then categ = "principal"
            if visibilidad = "" then visibilidad = 1   
        %>
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <div class="page-title-bar">
            <%= PageTitle %>
        </div>

        <form>
            <main>
                <div class="contenedor">
                    <br />

                    <div class="line">
                        <label>Ver Lista</label>
                        <select name="cboVisibilidad" id="cboVisibilidad">
                            <option value="1" <% if visibilidad = "1" then response.write " selected" %>>Contactos Activos</option>
                            <option value="0" <% if visibilidad = "0" then response.write " selected" %>>Contactos Obsoletos</option>
                            <option value="*" <% if visibilidad = "*" then response.write " selected" %>>Completa</option>
                        </select>                    
                    </div>

                    <div class="line">
                        <label>Tipos de Contactos</label>
                        <select name="cboTipo" id="cboTipo" onchange="RequeryCategs()">
                            <option value="*" <% if Tipo = "*" then response.write " selected='selected'" %>>- - Todos - -</option>
                            <%
                                sqlString ="SELECT Codigo, Nombre " & _
                                            "FROM con_Contactos_Tipos " & _
                                            "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                        "ORDER BY Nombre ASC;"

                                set cbox = cont_con.execute(sqlString)
                                    if not (cbox.bof or cbox.eof) then
                                        Do
                                            response.write "<option value='" & cbox("Codigo") & "' "
                                                if Tipo = cbox("Codigo") then 
                                                    response.write " selected='selected'" 
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
                        <label>Categorías</label>
                        <select name="cboCategoria" id="cboCategoria">
                            <option value="*" <% if Categ = "*" then response.write " selected='selected'" %>>- - Todos - -</option>
                            <%
                                sqlString = "SELECT Codigo, Nombre " & _
                                            "FROM con_Contactos_Categorias " & _
                                            "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                            "AND (Tipo = '" & Tipo & "') " & _
                                        "ORDER BY Nombre ASC;"

                                set cbox = cont_con.execute(sqlString)

                                if not (cbox.bof or cbox.eof) then
                                    Do
                                        response.write "<option value='" & cbox("Codigo") & "' "
                                            if categ = cbox("Codigo") then 
                                                response.write " selected='selected'" 
                                            end if
                                        response.write ">" & cbox("Nombre") & "</option>"

                                        cbox.MoveNext
                                    Loop Until cbox.eof
                                end if

                                cbox.close: set cbox = nothing
                            %>
                        </select>                    
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

            <button class="footer-button" type="button" aria-label="Aplicar" onclick="aplicar()">
                <i class="fa-solid fa-check"></i>
            </button>             
        </footer>

        <script>
            function volver() {
                history.back();
            }

            function irA(vinculo) {
                window.location.href = vinculo;
            }            

            function RequeryCategs() {
                guardarCookie("cont_v", document.getElementById("cboVisibilidad").value);
                guardarCookie("cont_t", document.getElementById("cboTipo").value);
                guardarCookie("cont_c", "*");

                var vinculo = "filtrar.asp";
                window.location.href = vinculo;
            }   
            
            function aplicar() {
                guardarCookie("cont_v", document.getElementById("cboVisibilidad").value);
                guardarCookie("cont_t", document.getElementById("cboTipo").value);
                guardarCookie("cont_c", document.getElementById("cboCategoria").value);

                var vinculo = "lista.asp";
                window.location.href = vinculo;
            }                
        </script>

        <% cont_con.close: set cont_con = nothing %>
        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>  