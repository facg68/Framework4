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
            thisProcess = "agenda.0090"
            SysLockOut


            dim pre_con, pre_t, pre_sqlString, pre_vinculo

            set pre_con = Server.CreateObject("ADODB.Connection")
            pre_con.open Application("Conn")        

            tipo = Request.Cookies("pre_t")
            estado = Request.Cookies("pre_e")

            if (isnull(tipo) or tipo = "")     then tipo = "P"
            if (isnull(estado) or estado = "") then estado = "1"
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
                        <label>Tipo</label>
                        <select name="cboTipo" id="cboTipo" onChange="Requery();">
                            <option value="P" <% if tipo = "P" then response.write " selected" %>>Presupuestos</option>
                            <option value="M" <% if tipo = "M" then response.write " selected" %>>Modelos</option>
                        </select>   
                    </div>

                    <div class="line">
                        <label>Estado</label>
                        <select name="cboEstatus" id="cboEstatus">
                            <option value="1" <% if estado = "1" then response.write " selected" %>>Abiertos</option>
                                <% if tipo = "P" then %>
                                    <option value="0" <% if estado = "0" then response.write " selected" %>>Cerrados</option>
                                <% end if %>
                            <option value="2" <% if estado = "2" then response.write " selected" %>>Archivados</option>
                        </select>
                    </div>

                    <br />
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

            function Requery() {
                guardarCookie("pre_t", document.getElementById("cboTipo").value);
                guardarCookie("pre_e", "1");

                var vinculo = "filtrar.asp";
                window.location.href = vinculo;
            }   
            
            function aplicar() {
                guardarCookie("pre_t", document.getElementById("cboTipo").value);
                guardarCookie("pre_e", document.getElementById("cboEstatus").value);

                var vinculo = "lista.asp";
                window.location.href = vinculo;
            }                
        </script>

        <% pre_con.close: set pre_con = nothing %>
        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>  