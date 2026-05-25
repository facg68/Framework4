<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Contactos" %>
        <title><%= PageTitle %></title>

        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0050"
            SysLockOut

            dim cont_con, cont_t, cont_sqlString, cont_vinculo
            dim visibilidad, tipo, categoria

            set cont_con= Server.CreateObject("ADODB.Connection")
            cont_con.open Application("Conn")

            function NombreMes(Mes)
                dim con_tt

                NombreMes = ""

                set con_tt = cont_con.Execute("SELECT Nombre FROM seg_cripto_Secuencias WHERE (Tipo = 'M') AND Valor = " & Mes & ";")
                    if not (con_tt.bof or con_tt.eof) then
                        NombreMes = con_tt("Nombre")
                    end if
                con_tt.close: set con_tt = nothing                        
            end function
        %>

        <style>
            .contacto-item {
                display: flex;
                align-items: center;
                gap: 14px;
                width: 100%;
                padding: 14px 12px;
                text-decoration: none;
                color: inherit;
                border-bottom: 1px solid rgba(0,0,0,0.08);
            }

            .contacto-item:active {
                background: rgba(0,0,0,0.06);
            }

            .contacto-foto {
                width: 52px;
                height: auto;              
                border-radius: 15px;       
                object-fit: contain;       
                background: transparent;
                display: block;
            }

            .contacto-info {
                flex: 1;
                min-width: 0;
            }

            .contacto-nombre {
                font-family: 'Ruda Bold';
                font-size: 1.15rem;                
                color: #111;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            }

            .contacto-telefono {
                font-family: 'Ruda';
                font-size: 1.05rem;
                color: #555;
            }            
        </style>        
    </head>

    <body plus=".contacto-item : plus_telCon ; ">
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <%
            visibilidad = Request.Cookies("cont_v")
            tipo = Request.Cookies("cont_t")
            categoria = Request.Cookies("cont_c")

            if visibilidad = "" then visibilidad = "1"
            if tipo = "" then tipo = "PE"
            if categoria = "" then categoria = "principal"

            cont_sqlString = "SELECT DISTINCT Usuario, Codigo, Nombre, Correo, Cumple, Telefono, NombreResumido " & _
                                    "FROM ( " & _
                                            " SELECT Usuario, Codigo, Nombre, Correo, Cumple, Telefono, TipoContacto, Categ, NombreResumido " & _
                                            " FROM con_FiltroContactos " & _
                                            " WHERE (Codigo <> '" & Request.Cookies("Usuario") & "') " & _
                                            " AND (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                            " AND (Telefono <> '') " 

            if visibilidad <> "*" then 
                cont_sqlString = cont_sqlString & " AND (Visible = " & visibilidad & ") " 
            end if

            if Tipo <> "*" then 
                cont_sqlString = cont_sqlString & " AND (TipoContacto = '" & tipo & "') "
            end if

            if categoria <> "*" then 
                cont_sqlString = cont_sqlString & " AND (Categ = '" & categoria & "') " 
            end if                        

            cont_sqlString = cont_sqlString & " ) as t ORDER BY Nombre ASC;"

            set cont_t = cont_con.Execute(cont_sqlString)        
        %>

        <div class="page-title-bar">
            <%= PageTitle %>
        </div>

        <main>
            <%
                if not (cont_t.bof or cont_t.eof) then
                    cuantos = 0

                    Do
                        cuantos = cuantos + 1

                        %>
                            <a class="contacto-item"
                               id="<%= cont_t("Codigo") %>"
                               data-nombre="<%= cont_t("NombreResumido") %>"
                               data-telefono="<%= Trim(cont_t("Telefono")) %>"
                               data-correo="<%= Trim(cont_t("correo"))%>"
                            >
                                <div>
                                    <img class="contacto-foto" src="<%= request.Cookies("usuPath") & "/fotos/" & cont_t("Codigo") & "_s.jpg" %>"
                                        onerror="this.src='/core/imagenes/misc/foto.jpg'">
                                </div>

                                <div class="contacto-info">
                                    <div class="contacto-nombre">
                                        <%= cont_t("Nombre") %>
                                    </div>

                                    <div class="contacto-telefono">
                                        <%
                                            response.write cont_t("Telefono") 

                                            if cont_t("correo") <> "" then 
                                                response.write "<br/>" & cont_t("correo")
                                            end if

                                            if cont_t("cumple") <> "" then 
                                                response.write "<br/>Cumpleaños: " & (left(cont_t("cumple"), 2) * 1) & " de " & NombreMes(right(cont_t("cumple"), 2))
                                            end if
                                        %>
                                    </div>
                                </div>
                            </a>
                        <%

                        cont_t.MoveNext
                    Loop Until cont_t.eof
                end if

                cont_t.close: set cont_t = nothing
            %>
        </main>

        <footer class="footer-contextual">
            <button class="footer-button" type="button" aria-label="Home" onclick="irA('/forma/mobile')">
                <i class="fa-solid fa-house"></i>
            </button>

            <button class="footer-button" type="button" aria-label="Volver" onclick="volver()">
                <i class="fas fa-undo-alt"></i>
            </button>

            <button class="footer-button" aria-label="Nuevo" onclick="nuevo()">
                <i class="fas fa-user-plus"></i>
            </button>            

            <button class="footer-button" aria-label="Filtrar" onclick="filtro()">
                <i class="fa-solid fa-list-check"></i>
            </button>            
        </footer>

        <script>
            function volver() {
                history.back();
            }

            function irA(vinculo) {
                window.location.href = vinculo;
            }
            
            function filtro() {
                var vinculo = "filtrar.asp";
                window.location.href = vinculo;
            }

            function nuevo() {
                var vinculo = "cont_nuevo.asp";
                window.location.href = vinculo;                
            }
        </script>

        <% cont_con.close: set cont_con = nothing %>
        
        <!-- #include virtual = "/forma/plus/telContactos.plus" -->     
        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>