<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Informe de Medios</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->  
        <%
            thisSystem = "discos"
            thisProcess = "discos.0130"
            SysLockOut

            dim cc, tt, sqlString, Usuario, Editor, ckIndice, ckNombre, ckValor

            Usuario = Request.Cookies("usuario")
            Editor = Request.QueryString("e")
            
            if (Editor = "") then Editor = "DM"

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")            
        %>   

		<style>
            p {
                margin-bottom: 10px;
                line-height: 1.5;
            }  
		</style>            
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->    
        <br />

        <form name="form_transaccion" id="form_transaccion" method="post" action="rep_medios_rep.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 35%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Informe de Medios
                </div>
                
                <div style="flex: 0 0 65%; text-align: right;">
                    <select class="no-field" name="cboEditor" id="cboEditor" onChange="Refrescar(1)" >
                        <%
                            set tt = cc.execute("SELECT Codigo, Nombre FROM discos_Objetos_Clases ORDER BY Nombre;")

                            if not (tt.bof or tt.eof) then
                                do
                                    response.write "<option value='" & tt("Codigo") & "' "
                                        if Editor = tt("Codigo") then 
                                            response.write " selected" 
                                        end if
                                    response.write ">" & tt("Nombre") & "&nbsp;&nbsp;&nbsp;&nbsp;</option>"

                                    tt.MoveNext
                                loop until tt.eof
                            end if

                            tt.close: set tt = nothing
                        %>
                    </select>  

                    &nbsp;                
                    
                    <button class="form-btn verde normal" type="button" onclick="informe()">Informe</button>                
                </div>

                <div class="no-ver">
                    <input type="text" id="frm_Nivel" name="frm_Nivel" value="<%= Nivel %>">
                </div>
            </div>   

            <div class="main main-scroll">
                <div class="no-ver">
                </div>

                <div class="line">
                    <label class="label normal">Forma</label>
                    <div class="label full section">
                        <%
                            sqlString = "SELECT f.Usuario, f.Forma, f.Nombre, f.Icono_Forma " & _
                                        "FROM discos_Formas AS f " & _
                                        "INNER JOIN ( " & _
                                                    "SELECT DISTINCT Usuario, Forma " & _
                                                    "FROM discos_Objetos " & _ 
                                                    "WHERE (Usuario = '" & usuario & "') " & _
                                                    "AND (Editor = '" & Editor & "') " & _
                                                ") AS q " & _
                                        "ON f.Forma = q.Forma " & _
                                        "AND f.Usuario = q.Usuario " & _ 
                                        "ORDER BY f.Nombre;" 

                            set tt = cc.execute(sqlString)

                            if not (tt.bof or tt.eof) then
                                ckIndice = 0

                                do
                                    ckNombre = "f" & ckIndice
                                    response.write "<p><input type='checkbox' id='" & ckNombre & "' name='" & ckNombre & "' value='" & tt("Forma") & "'"
                                        if Request.QueryString(ckNombre) <> "" then response.write " checked"
                                    response.write " onChange='Refrescar(2)'>&nbsp;"
                                    response.write "<label>" & tt("Nombre") & "</label></p>"

                                    ckIndice = ckIndice + 1
                                    tt.MoveNext
                                loop until tt.eof
                            end if

                            tt.close: set tt = nothing
                        %>

                        <div class="no-ver">    
                            <input type="text" id="frm_Formas" name="frm_Formas" value="<%= ckIndice %>">
                        </div>                    
                    </div>
                </div>

                <%
                    if (Editor = "JU") OR (Editor = "SO") then 
                        clase = "line"
                    else
                        clase = "no-ver"
                    end if
                %>

                <div class="<%= clase %>">
                    <label class="label normal">Plataforma</label>
                    <div class="label full section">
                        <% 
                            if (Editor = "JU") OR (Editor = "SO") then 
                                ckIndice = Request.QueryString("numFormas")
                                inLista = ""

                                for k = 0 to (ckIndice - 1)
                                    ckNombre = "f" & k

                                    if (Request.QueryString(ckNombre) <> "") OR (Request.QueryString(ckNombre) <> NULL)  then
                                        if len(trim(inLista)) = 0 then
                                            inLista = "'" & Request.QueryString(ckNombre) & "'"
                                        else
                                            inLista = inLista & ", '" & Request.QueryString(ckNombre) & "'"
                                        end if
                                    end if
                                next

                                if inLista = "" then inLista = "'*'"

                                sqlString = "SELECT p.Usuario, p.Codigo, p.Nombre " & _
                                            "FROM discos_Plataformas AS p " & _
                                            "INNER JOIN ( " & _
                                                            "SELECT DISTINCT Usuario, PlatOS " & _
                                                            "FROM discos_Objetos " & _
                                                            "WHERE (Usuario = '" & Usuario & "') " & _
                                                            "AND (Editor = '" & Editor & "') " & _
                                                            "AND (Forma IN (" & inLista & ")) " & _
                                                        ") as q " & _
                                            "ON (p.Usuario = q.Usuario) " & _
                                            "AND (p.Codigo = q.PlatOS) " & _
                                            "AND (p.Codigo <> '00000000') " & _
                                            "ORDER BY p.Nombre;"

                                set tt = cc.execute(sqlString)

                                if not (tt.bof or tt.eof) then
                                    ckIndice = 0

                                    do
                                        ckNombre = "p" & ckIndice
                                        response.write "<p><input type='checkbox' id='" & ckNombre & "' name='" & ckNombre & "' value='" & tt("Codigo") & "'"
                                            if Request.QueryString(ckNombre) <> "" then response.write " checked"
                                        response.write " onChange='Refrescar(3)'>&nbsp;"
                                        response.write "<label>" & tt("Nombre") & "</label></p>"

                                        ckIndice = ckIndice + 1
                                        tt.MoveNext
                                    loop until tt.eof
                                end if

                                tt.close: set tt = nothing
                            end if
                        %>              

                        <div class="no-ver">
                            <input type="text" id="frm_Plataformas" name="frm_Plataformas" value="<%= ckIndice %>">
                        </div>                    
                    </div>
                </div>

                <div class="line">
                    <label class="label normal">Opciones</label>
                    <div class="label full section">
                        &nbsp;

                        <input type="radio" id="chk_ruptura" name="chk_ruptura" value="0" checked="checked">
                        <label for="chk_rupturaPlat">&nbsp;No Agrupar</label><br />

                        <% if Editor <> "HW" then %>
                            <br />&nbsp;
                            <input type="radio" id="chk_ruptura" name="chk_ruptura" value="1">
                            <label for="chk_rupturaPlat">&nbsp;Agrupar Plataformas</label><br />

                            <br />&nbsp;
                            <input type="radio" id="chk_ruptura" name="chk_ruptura" value="2">
                            <label for="chk_ordenMeta">&nbsp;Agrupar Metadata</label>

                            <br />                
                        <% end if %>                      
                    </div>
                </div>
            </div>
        </form>

        <br /><br />

        <script>
            function informe() {
                document.getElementById("form_transaccion").submit(); 
            }

            function Refrescar(nivel) {
                var editor = document.getElementById("cboEditor").value;
                var nomForma = "";

                if (nivel == 1) {
                    var vinculo = "rep_medios.asp?e=" + editor;
                    window.location.href = vinculo;
                };

                if (nivel == 2) {
                    var vinculo = "rep_medios.asp?e=" + editor;
                    var formas = document.getElementById("frm_Formas").value;         

                    vinculo += "&numFormas=" + (parseInt(formas));

                    for (let i = 0; i < formas; i++) { 
                        nomForma = "f" + i;

                        if (document.getElementById(nomForma).checked) {
                            vinculo += "&" + nomForma + "=" + document.getElementById(nomForma).value;
                        } else {
                            vinculo += "&" + nomForma + "=";
                        };           
                    };

                    window.location.href = vinculo;
                };          

                if (nivel == 3) {
                    var vinculo = "rep_medios.asp?e=" + editor;
                    var formas = document.getElementById("frm_Formas").value;              
                    var plataformas = document.getElementById("frm_Plataformas").value;

                    vinculo += "&numFormas=" + (parseInt(formas));

                    for (let i = 0; i < formas; i++) { 
                        nomForma = "f" + i;

                        if (document.getElementById(nomForma).checked) {
                            vinculo += "&" + nomForma + "=" + document.getElementById(nomForma).value;
                        } else {
                            vinculo += "&" + nomForma + "=";
                        };           
                    };

                    vinculo += "&numPlataformas=" + (parseInt(plataformas));            

                    for (let i = 0; i < plataformas; i++) {            
                        nomForma = "p" + i;

                        if (document.getElementById(nomForma).checked) {
                            vinculo += "&" + nomForma + "=" + document.getElementById(nomForma).value;
                        } else {
                            vinculo += "&" + nomForma + "=";
                        };           
                    };  

                    window.location.href = vinculo;        
                };
            }    
        </script>

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->    
        <% cc.close: set cc = nothing %>    
    </body>
</html>