<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <title>Buscar</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" --> 

        <%
            dim con, t, sqlString, imgTipo
            dim vinculo, sw, Cadena
            dim ckCon, ckCal, ckPre, ckPk, ckObj, ckDet, ckPro, ckMeta

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            ' Funciones y Procedimientos ------------------------------------------------
                Function SiNull(Campo, ValorReemplazo)
                    if (isnull(Campo)) Or (Len(trim(Campo)) = 0) then
                        siNull = ValorReemplazo
                    else
                        siNull = 1
                    end if
                end function  

                Function GenerarWhere(cadena, campo)
                    Dim gruposOR, grupo, partesAND, parte
                    Dim whereFinal, clausulaOR, clausulaAND
                    Dim i, j

                    cadena = Trim(cadena)

                    If cadena = "" Then
                        GenerarWhere = ""
                        Exit Function
                    End If

                    gruposOR = Split(cadena, ",")

                    For i = 0 To UBound(gruposOR)
                        grupo = Trim(gruposOR(i))
                        partesAND = Split(grupo, "+")

                        clausulaAND = ""

                        For j = 0 To UBound(partesAND)
                            parte = Trim(partesAND(j))
                            ' Escapar comillas simples en caso de que el usuario haya escrito algo como O'Brien
                            parte = Replace(parte, "'", "''")
                            If clausulaAND <> "" Then clausulaAND = clausulaAND & " AND "
                            clausulaAND = clausulaAND & "(" & campo & " LIKE ''%" & parte & "%'')"
                        Next

                        If clausulaOR <> "" Then clausulaOR = clausulaOR & " OR "
                        clausulaOR = clausulaOR & "(" & clausulaAND & ")"
                    Next

                    GenerarWhere = clausulaOR
                End Function      
            ' Fin - Funciones y Procedimientos ------------------------------------------
        %>  

        <style>
            .field {
                padding: 0.3rem 0.4rem;
                border: 1px solid #cccccc;
                border-radius: 0.3rem;
                font-family: 'Ruda', sans-serif;
                font-size: 1rem;
                color: rgb(25, 25, 25);                
                background-color:  rgb(245, 235, 250);
                box-sizing: border-box;
                resize: vertical;
            }    
            
            .busqueda-label {
                font-size: 18px;
                font-family: 'Ruda', sans-serif;
                margin-right: 10px;
                color: #333; /* o el color que uses en tu theme */
            }  
            
            .chk-line {
                display: inline-flex;
                align-items: center;
                margin-right: 15px; /* separa cada grupo */
            }

            .chk-line input[type="checkbox"] {
                margin: 0 5px 0 0;
                vertical-align: middle;
            }            
        </style>          
    </head>

    <body plantilla="tabla" reserva="185">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            ckCon  = SiNull(Request.Form("ckCon"),  0)
            ckCal  = SiNull(Request.Form("ckCal"),  0)
            ckPre  = SiNull(Request.Form("ckPre"),  0)
            ckPk   = SiNull(Request.Form("ckPk"),   0)
            ckObj  = SiNull(Request.Form("ckObj"),  0)
            ckDet  = SiNull(Request.Form("ckDet"),  0)
            ckPro  = SiNull(Request.Form("ckPro"),  0)
            ckMeta = SiNull(Request.Form("ckMeta"), 0)

            cuantos = 0
            Usuario = Request.Cookies("Usuario")
            Cadena = Request.Form("txtCadena")    
        %>

        <br />

        <form id="formulario" name="formulario" method="post" action="buscar.asp" class="cmxform form-horizontal style-form">
            <div style="display: flex; flex-direction: column; width: 95%; margin: auto; gap: 10px;">
                <div style="display: flex; align-items: center; justify-content: space-between;">
                    <div style="flex: 0 0 auto;">
                        <label for="txtCadena" class="busqueda-label">Buscar</label>
                    </div>

                    <input class="field" style="flex: 1 1 auto; margin: 0 10px;"
                        id="txtCadena" name="txtCadena" 
                        type="text" value="<%= Cadena %>"
                        placeholder="Buscar..."
                    >

                    <button class="form-btn verde violeta" type="button" onclick="Buscar();">
                        <i class="fa-solid fa-magnifying-glass"></i>
                    </button>
                </div>

                <div style="display: flex; flex-wrap: wrap; gap: 15px; justify-content: center;">
                    <div class="chk-line">
                        <input type="checkbox" id="ckCon" name="ckCon" <% if ckCon = 1 then response.write " checked" %> onChange="VerifCheck()">
                        <label for="chk1">Gente</label>
                    </div>

                    <div class="chk-line">
                        <input type="checkbox" id="ckCal" name="ckCal" <% if ckCal = 1 then response.write " checked" %> onChange="VerifCheck()">
                        <label for="chk1">Eventos</label>
                    </div>

                    <div class="chk-line">
                        <input type="checkbox" id="ckPre" name="ckPre" <% if ckPre = 1 then response.write " checked" %> onChange="VerifCheck()">
                        <label for="chk1">Gastos</label>
                    </div>

                    <div class="chk-line">
                        <input type="checkbox" id="ckPk" name="ckPk" <% if ckPk = 1 then response.write " checked" %> onChange="VerifCheck()">
                        <label for="chk1">Medios</label>
                    </div>

                    <div class="chk-line">
                        <input type="checkbox" id="ckObj" name="ckObj" <% if ckObj = 1 then response.write " checked" %> onChange="VerifCheck()">
                        <label for="chk1">Objetos</label>
                    </div>

                    <div class="chk-line">
                        <input type="checkbox" id="ckDet" name="ckDet" <% if ckDet = 1 then response.write " checked" %> onChange="VerifCheck()">
                        <label for="chk1">Detalles</label>
                    </div>

                    <div class="chk-line">
                        <input type="checkbox" id="ckPro" name="ckPro" <% if ckPro = 1 then response.write " checked" %> onChange="VerifCheck()">
                        <label for="chk1">Actores</label>
                    </div>

                    <div class="chk-line">
                        <input type="checkbox" id="ckMeta" name="ckMeta" <% if ckMeta = 1 then response.write " checked" %> onChange="VerifCheck()">
                        <label for="chk1">Metadata</label>
                    </div>
                </div>
            </div>
        </form>

        <div class="main main-scroll" style="width: 95%;">
            <div class="line">
                <div class="tabla-wrapper">
                    <% 
                        If Cadena <> "" then
                            con.execute ("DELETE FROM seg_SuperBusquedas WHERE Usuario = '" & Usuario & "';")

                            if ckCal  = 1 then con.execute("seg_BuscarCalendarios '" & Usuario & "', '" & Cadena & "';")

                            if ckCon  = 1 then 
                                con.execute("seg_BuscarContactos_01 '" & Usuario & "', '" & Cadena & "';")
                                con.execute("seg_BuscarContactos_02 '" & Usuario & "', '" & Cadena & "';")
                                con.execute("seg_BuscarContactos_03 '" & Usuario & "', '" & Cadena & "';")
                            end if

                            if ckPre  = 1 then con.execute("seg_BuscarPresupuestos '" & Usuario & "', '" & Cadena & "';")
                            if ckPk   = 1 then con.execute("seg_BuscarDiscosPaquetes '" & Usuario & "', '" & Cadena & "';")
                            if ckObj  = 1 then con.execute("seg_BuscarDiscosObjetos '" & Usuario & "', '" & Cadena & "';")
                            if ckDet  = 1 then con.execute("seg_BuscarDiscosDetalles '" & Usuario & "', '" & Cadena & "';")
                            if ckPro  = 1 then con.execute("seg_BuscarDiscosProtagonista '" & Usuario & "', '" & GenerarWhere(Cadena, "Protagonista") & "';")
                            if ckMeta = 1 then con.execute("seg_BuscarDiscosMetadata '" & Usuario & "', '" & GenerarWhere(Cadena, "Lista_Metadata") & "';")

                            set t = con.execute("SELECT * FROM seg_SuperBusquedas WHERE Usuario = '" & Usuario & "' ORDER BY ISNULL(AEdicion, 0), Titulo; ")
                                %>
                                    <table class="tabla tabla-violet">
                                        <thead>
                                            <tr>
                                                <th class="sticky" style="width: 10%; text-align: center;">Clase</th>
                                                <th class="sticky" style="width: 10%; text-align: center;">Imagen</th>
                                                <th class="sticky" style="width: 70%; text-align: center;">Titulo</th>
                                                <th class="sticky" style="width: 10%; text-align: center;">Fecha</th>
                                            </tr>
                                        </thead>   

                                        <%
                                            if not (t.bof or t.eof) then
                                                response.write "<tbody>"

                                                Do
                                                    cuantos = cuantos + 1

                                                    select case t("tipo")
                                                        case "cal"
                                                            vinculo = "/forma/desktop/apps/agenda/cal/cal_eventos.asp?s=" & t("Llave1")
                                                            imgTipo = "/core/imagenes/misc/icono_calendario.png"
                                                        case "pre"
                                                            vinculo = "/forma/desktop/apps/agenda/pre/presupuestos/pre_det_tra_editar.asp?p=" & t("Llave1") & "&l=" & t("Llave2")
                                                            imgTipo = "/core/imagenes/misc/icono_presupuesto.png"
                                                        case "con"
                                                            vinculo = "/forma/desktop/apps/agenda/cont/cont_editar.asp?con=" & t("Llave1")
                                                            imgTipo = "/core/imagenes/misc/icono_gente.png"
                                                        case "dpk", "dobj"
                                                            vinculo = "/forma/desktop/apps/discos/inventario/medios/editar.asp?m=" & t("Llave1")
                                                            imgTipo = "/core/imagenes/misc/icono_libreria.png"
                                                        case "ddet"
                                                            vinculo = "/forma/desktop/apps/discos/inventario/medios/editar_objeto.asp?p=" & t("llave1") & "&o=" & t("Llave2") & "&e=" & t("Llave3")
                                                            imgTipo = "/core/imagenes/misc/icono_libreria.png"
                                                    end select

                                                    %>
                                                        <tr onclick="AbrirVinculo('<%= vinculo %>')"> 
                                                            <td style="text-align: center;">
                                                                <img src="<%= imgTipo %>" width="80">
                                                            </td>

                                                            <td style="text-align: center;">
                                                                <%
                                                                    Select Case t("Tipo") 
                                                                        case "con"
                                                                            fotoPaquete = "/perfiles/" & Usuario & "/fotos/" & t("LLave1") & "_s.jpg"
                                                                            fotoObjeto = "/core/imagenes/misc/foto.jpg"
                                                                            %><img src="<%= fotoPaquete %>" onerror="this.src='<%= fotoObjeto %>'" width="80"><%

                                                                        case "pre"
                                                                            if t("Llave3") <> "" then
                                                                                fotoPaquete = "/perfiles/" & Usuario & "/fotos/" & t("LLave3") & "_s.jpg"
                                                                                fotoObjeto = "/core/imagenes/misc/foto.jpg"
                                                                                %><img src="<%= fotoPaquete %>" onerror="this.src='<%= fotoObjeto %>'" width="80"><%
                                                                            else
                                                                                response.write "&nbsp;"
                                                                            end if

                                                                        case "dpk"
                                                                                fotoPaquete = "/perfiles/" & Usuario & "/medios/" & t("LLave1") & "_s.jpg"
                                                                                fotoObjeto = "/core/imagenes/misc/foto.jpg"
                                                                                %><img src="<%= fotoPaquete %>" onerror="this.src='<%= fotoObjeto %>'" width="80"><%

                                                                        case "dobj", "ddet"                     
                                                                                fotoPaquete = "/perfiles/" & Usuario & "/medios/" & t("LLave1") & "_s.jpg"
                                                                                fotoObjeto = "/perfiles/" & Usuario & "/medios/" & t("LLave2") & "_s.jpg"
                                                                                %><img src="<%= fotoObjeto %>" onerror="this.src='<%= fotoPaquete %>'" width="80"><%

                                                                        case else
                                                                            response.write "&nbsp;"

                                                                    end select
                                                                %>
                                                            </td>

                                                            <td style="text-align: left;">
                                                                <%
                                                                    select case t("Tipo")
                                                                        case "con"
                                                                            response.write "<b>" & t("Titulo") & "</b><br/>"
                                                                            response.write "Telefono: " & t("Telefono") & "<br/>"                    
                                                                            response.write "Pais: " & t("Pais") & "<br/>"
                                                                            response.write "Correo: " & t("Correo") & "<br/>"
                                                                            response.write "Stio Web: " & t("Web") & "<br/>"

                                                                        case "cal", "pre"
                                                                            response.write "<b>" & t("Titulo") & "</b><br/>"
                                                                            response.write t("Descripcion") & "<br/>"
                                                                            response.write t("Ubicacion") & "<br/>"
                                                                            response.write t("Contacto") & "<br/>"
                                                                            response.write t("Monto") & "<br/>"

                                                                        case "dpk"
                                                                            response.write "<b>" & t("Titulo") & "</b><br/>"
                                                                            response.write "Año de Edición: " & t("AEdicion") & ", Año de Compra: " & t("ACompra") & "<br/>"
                                                                            response.write "Tienda: " & t("Empresa") & "<br/>"
                                                                            response.write "Precio: " & t("Monto") & "<br/>"
                                                                            response.write "Tipo: Paquete<br/>"

                                                                        case "dobj"
                                                                            response.write "<b>" & t("Titulo") & "</b><br/>"
                                                                            response.write "Año de Edición: " & t("AEdicion") & ", Año de Compra: " & t("ACompra") & "<br/>"
                                                                            response.write "Tienda:" & t("Empresa") & "<br/>"
                                                                            response.write "Precio: " & t("Monto") & "<br/>"
                                                                            response.write "Tipo: Objeto<br/>"

                                                                        case "ddet"
                                                                            response.write "<b>" & t("Titulo") & "</b><br/>"
                                                                            response.write "Año de Edición: " & t("AEdicion") & ", Año de Compra: " & t("ACompra") & "<br/>"                                                    
                                                                            response.write "Tienda:" & t("Empresa") & "<br/>"
                                                                            response.write "Precio: " & t("Monto") & "<br/>"
                                                                            response.write "Tipo: Detalle<br/>"
                                                                    end select
                                                                %>
                                                            </td>

                                                            <td style="text-align: center;">
                                                                <%
                                                                    select case t("Tipo")
                                                                        case "cal", "pre"
                                                                            response.write Replace(t("FechaHora"), " ", "<br/>")
                                                                        case else   
                                                                            response.write "&nbsp;"
                                                                    end select
                                                                %>
                                                            </td>
                                                        </tr>
                                                    <%

                                                    t.MoveNext
                                                Loop Until t.eof

                                                response.write "</tbody>"
                                            end if
                                        %>

                                        <tfoot>
                                            <tr>
                                                <td colspan="4" class="sticky" style="text-align: center;">
                                                    <%
                                                        Select Case cuantos
                                                            case 0: response.write "No hay Resultados"
                                                            case 1: response.write "Un Resultado"
                                                            case else
                                                                response.write cuantos & " resultados"
                                                        end Select
                                                    %>
                                                </td>
                                            </tr>
                                    </table>   
                                <%
                            t.close: set t = nothing
                        end if                                    
                    %> 
                </div>
            </div>
        </div>
  
        <br /><br />   

        <script type="text/javascript">
            function irA(vinculo) {
                window.location.href =vinculo;
            }

            function VerifCheck() {
                const checkboxes = ["ckCon", "ckCal", "ckPre", "ckPk", "ckObj", "ckDet", "ckPro", "ckMeta"];
                return checkboxes.some(id => document.getElementById(id)?.checked);
            }

            function NoTieneValor(value) {
                return value === null || value === undefined || value.toString().trim() === "";
            }            

            function AbrirVinculo(vinculo) {
                window.location.href = vinculo;
            }

            function Buscar() {
                if (NoTieneValor(document.getElementById("txtCadena").value) == true) {
                    alert("Debe escribir un valor en el cuadro de texto antes de iniciar una consulta");
                } else {
                    if (VerifCheck() == false) {
                        document.getElementById("ckPk").checked = true;
                    };

                    document.getElementById("formulario").submit();             
                } 
            }

            document.getElementById("txtCadena").addEventListener("keydown", function(event) {
                if (event.key === "Enter") {
                    event.preventDefault(); // Previene que se envíe un formulario, si aplica
                    Buscar();
                }
            });            
        </script>   

        <% con.close: set con = nothing %> 
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->        
    </body>