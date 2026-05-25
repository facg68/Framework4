<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Filtrar Presupuestos" %>
        <title><%= PageTitle %></title>

        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0090"
            SysLockOut

            dim pre_con, pre_t, pre_sqlString, pre_vinculo

            set pre_con = Server.CreateObject("ADODB.Connection")
            pre_con.open Application("Conn")        

            usu = Request.Cookies("Usuario")
            pre = Request.Cookies("pre_det_pre")
            dia = Request.Cookies("pre_det_dia")
            ver = Request.Cookies("pre_det_ver")

            if dia = "" then dia = "*"
            if ver = "" then ver = "0"
        %>
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <div class="page-title-bar">
            <div class="title"><%= PageTitle %></div>
        </div>

        <form>
            <main>
                <div class="contenedor">
                    <br />

                    <div class="line">
                        <label>Estado</label>

                        <select name="cboVer" id="cboVer" onchange="recargarMostrar();">
                            <option value="0" <% if ver = "0" then response.write " selected" %>>Ver Pendientes</option>
                            <option value="1" <% if ver = "1" then response.write " selected" %>>Ver Cerradas</option>     
                            <option value="*" <% if ver = "*" then response.write " selected" %>>Ver Todas</option>                                                 
                        </select>  
                    </div>

                    <div class="line">
                        <label>Día</label>
                        <%
                            sqlString = "SELECT DISTINCT TOP (100) PERCENT Fecha, dbo.Cripto_FechaFormulario(Fecha) AS FF " & _
                                        "FROM dbo.pre_Presupuesto_Detalles AS p " & _
                                        "WHERE (Presupuesto = '" & pre & "') " & _
                                        "AND (Usuario = '" & usu & "') " 

                            if ver <> "*" then sqlString = sqlString & "AND (Aplicado = " & ver & ") "

                            sqlString = sqlString & "ORDER BY Fecha DESC;"

                            set l = pre_con.Execute(sqlString)                       
                        %>
                            <select name="cboDia" id="cboDia" onchange="recargarMostrar();">
                                <option value="*" <% if dia = "*" then response.write " selected" %>>- - Todos - -</option>
                                <%
                                    if not (l.bof or l.eof) then
                                        do
                                            response.write "<option value='" & l("Fecha") & "' " 
                                                if dia = l("Fecha") then 
                                                    response.write " selected"
                                                end if
                                            response.write ">" & left(l("FF") , 10) & "</option>"
                                        
                                            l.MoveNext
                                        loop until l.eof
                                    end if
                                %>
                            </select>
                            
                        <% l.close: set l = nothing %>
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

            function recargarMostrar() {
                guardarCookie("pre_det_dia", document.getElementById("cboDia").value);
                guardarCookie("pre_det_ver", document.getElementById("cboVer").value);

                var vinculo = "filtrar.asp";
                window.location.href = vinculo;
            }   
            
            function aplicar() {
                guardarCookie("pre_det_dia", document.getElementById("cboDia").value);
                guardarCookie("pre_det_ver", document.getElementById("cboVer").value);

                var vinculo = "pre_det_editar.asp";
                window.location.href = vinculo;
            }                
        </script>

        <% pre_con.close: set pre_con = nothing %>
        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>  