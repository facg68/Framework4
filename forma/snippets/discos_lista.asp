<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    Snip_Width = 1000 
%>

<!-- #include virtual = "/core/includes/snippets.inc" -->

<style>
    a.discos_lista_linea:link,
    a.discos_lista_linea:visited,
    a.discos_lista_linea:focus,
    a.discos_lista_linea:hover,
    a.discos_lista_linea:active {
        color: black !important;
    }
</style>

<%
    set discos_lista_con = Server.CreateObject("ADODB.Connection")
    discos_lista_con.open Application("Conn")


    function discos_lista_CarpetaPorDefecto()
        dim cc, tt

        sqlString = "SELECT Codigo FROM discos_Carpetas WHERE Usuario = '" & Request.Cookies("Usuario") & "' AND PorDefecto = 1;"

        set cc = Server.CreateObject("ADODB.Connection")
        cc.open Application("Conn")
            set tt = cc.execute(sqlString)
                discos_lista_CarpetaPorDefecto = tt("Codigo")
            tt.close: set tt = nothing
        cc.close: set cc = nothing
    end function    


    '
    ' Abrimos la tabla y llenamos los datos
    '
    dim discos_lista_con, usu, folder, tipo, forma, plataforma, amo
    
    usu = Request.Cookies("Usuario")
    folder = discos_lista_CarpetaPorDefecto()
%>

<!-- FILTROS -->
    <form>
        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 100%; text-align: center; font-size: 25px; color: rgb(50, 50, 50);">
                <select class="no-field" name="cboFolder" id="cboFolder" onChange="discos_lista_Requery('cboFolder');">
                    <%
                        sqlString = "Select Codigo, Nombre from discos_Carpetas WHERE Usuario = '" & Request.Cookies("Usuario") & "' ORDER BY Nombre;"
                        set tt = discos_lista_con.execute(sqlString)
                            if not (tt.bof or tt.eof) then
                                Do
                                    response.write "<option value='" & tt("Codigo") & "'"
                                        if tt("Codigo") = folder then
                                            response.write " selected"
                                        end if
                                    response.write ">" & tt("Nombre") & "</option>"

                                    tt.MoveNext
                                Loop Until tt.eof
                            end if
                        tt.close: set tt = nothing
                    %>
                </select>          

                <select class="no-field" name="cboTipo" id="cboTipo" onChange="discos_lista_Requery('cboTipo');">
                    <option value="*">- - Tipo - -</option>
                    <%
                        set tt = discos_lista_con.execute("select Codigo, Nombre from discos_Objetos_Clases ORDER BY Nombre ")
                            if not (tt.bof or tt.eof) then
                                Do
                                    response.write "<option value='" & tt("Codigo") & "'>" & tt("Nombre") & "</option>"
                                    tt.MoveNext
                                Loop Until tt.eof
                            end if
                        tt.close: set tt = nothing
                    %>
                </select>   

                <select class="no-field" name="cboPlataforma" id="cboPlataforma" onChange="discos_lista_Requery('cboPlataforma');">
                    <option value="*">- - Plataforma - -</option>
                        <%
                            lsqlString = "select Codigo, Nombre " & _
                                            "from discos_Plataformas " & _ 
                                            "where (usuario = '" & Request.Cookies("Usuario") & "') " & _
                                            "and (Codigo <> '00000000') " 

                            SELECT CASE Tipo
                                CASE "JU"
                                    lsqlString = lsqlString & "and (Juegos = 1) "
                                CASE "SO"
                                    lsqlString = lsqlString & "and (Software = 1) "
                            END SELECT

                            lsqlString = lsqlString & "order by Nombre"

                            set tt = discos_lista_con.execute(lsqlString)
                                if not (tt.bof or tt.eof) then
                                    Do
                                        response.write "<option value='" & tt("Codigo") & "'>" & tt("Nombre") & "</option>"
                                        tt.MoveNext
                                    Loop Until tt.eof
                                end if
                            tt.close: set tt = nothing
                        %>
                </select>   

                <select  class="no-field" name="cboForma" id="cboForma" onChange="discos_lista_Requery('cboForma');">
                    <option value="*">- - Forma - -</option>
                    <%
                        sqlString = "select Forma, Nombre " & _
                                      "from discos_Formas " & _
                                     "where Usuario = '" & Request.Cookies("Usuario") & "' " 
                        
                        Select Case Tipo 
                            Case "DM", "VM": sqlString = SqlString & "and Musica = 1 "
                            Case "PE": sqlString = SqlString & "and Video = 1 "
                            Case "JU": sqlString = SqlString & "and Juegos = 1 "
                            Case "SO": sqlString = SqlString & "and Software = 1 "
                            Case "LI": sqlString = SqlString & "and Libros = 1 "
                            Case "HW": sqlString = SqlString & "and Hardware = 1 "
                        end Select

                        sqlString = sqlString & "order by Nombre "

                        set tt = discos_lista_con.execute(sqlString)

                        if not (tt.bof or tt.eof) then
                            Do
                                response.write "<option value='" & tt("Forma") & "'>" & tt("Nombre") & "</option>"
                                tt.MoveNext
                            Loop Until tt.eof
                        end if

                        tt.close: set tt = nothing
                    %>
                </select>      

                <select class="no-field" name="txtAmo" id="txtAmo" onChange="discos_lista_Requery('txtAmo');">
                    <optgroup label="Año de Edición">
                        <option value="1"><%= Year(Date) %></option>
                        <option value="2"><%= (Year(Date) - 1) %></option>
                        <option value="3">Ultimos 2 Años</option>
                        <option value="4">Ultimos 5 Años</option>
                        <option value="5">Ultimos 10 Años</option>
                        <option value="6">Ultimos 15 Años</option>
                        <option value="7">Ultimos 20 Años</option>
                    </optgroup>

                    <optgroup label="Año de Compra">
                        <option value="8" ><%= Year(Date) %></option>
                        <option value="9" ><%= (Year(Date) - 1) %></option>
                        <option value="10" selected>Ultimos 2 Años</option>
                        <option value="11">Ultimos 5 Años</option>
                        <option value="12">Ultimos 10 Años</option>
                        <option value="13">Ultimos 15 Años</option>
                        <option value="14">Ultimos 20 Años</option>
                    </optgroup>

                    <optgroup label="- - - - - - - - - - - -">
                        <option value="0">Ver Todo</option>
                    </optgroup>
                </select>  
            
                <select class="no-field" name="cboOrden" id="cboOrden" onChange="discos_lista_Requery('cboOrden');">
                    <option value="1" >▲ A&ntilde;o</option>
                    <option value="2" >▲ Titulo</option>
                    <option value="3" >▲ Casa</option>
                    <option value="4" >▲ Medios</option>
                    <option value="5" >▲ Precio</option>

                    <option value="6" selected>▼ A&ntilde;o</option>
                    <option value="7" >▼ Titulo</option>
                    <option value="8" >▼ Casa</option>
                    <option value="9" >▼ Medios</option>
                    <option value="10">▼ Precio</option>            
                </select>           
            </div>
        </div>    
    </form>
<!-- FIN FILTROS -->    

<!-- TABLA CON DATOS -->
    <div class="main" style="max-height: 500px;">
        <div class="tabla-wrapper" id="discos_lista_body">
            <div style="text-align: center;">Cargando Datos...</div>
        </div>
    </div>

    <br />    
<!-- FIN TABLA CON DATOS -->

<script>
    function discos_lista_init() {
        redimWindow("discos_lista", <%= Snip_Width %>)

        const cbo = document.getElementById("cboPlataforma");
        if (cbo) {
            cbo.style.display = "none";
            discos_lista_Requery("cboFolder");
        }
    }

    function discos_plataformas_get() {
        return [
            <%
                sql = "SELECT Codigo, Nombre, Juegos, Software " & _
                        "FROM discos_Plataformas " & _
                       "WHERE Usuario = '" & Request.Cookies("Usuario") & "' " & _
                         "AND Obsoleta = 0 " & _
                    "ORDER BY Nombre"

                set tt = discos_lista_con.execute(sql)
                    primero = true

                    do while not tt.eof
                        if not primero then response.write ","
                        primero = false

                        response.write "{"
                        response.write "codigo:'" & tt("Codigo") & "',"
                        response.write "nombre:'" & Replace(tt("Nombre"),"'","\'") & "',"
                        response.write "juegos:" & tt("Juegos") & ","
                        response.write "software:" & tt("Software")
                        response.write "}"

                        tt.movenext
                    loop
                tt.close: set tt = nothing
            %>
        ];
    }

    function discos_lista_Requery(origen) {
        const folder = document.getElementById("cboFolder").value;
        const tipo   = document.getElementById("cboTipo").value;
        const forma  = document.getElementById("cboForma").value;
        const orden  = document.getElementById("cboOrden").value;
        const amo    = document.getElementById("txtAmo").value;

        const plataformaControl = document.getElementById("cboPlataforma");

        if (origen === "cboTipo") {
            discos_lista_cargarFormas(tipo);

            if (tipo === "JU" || tipo === "SO") {
                discos_lista_cargarPlataformas(tipo);
                plataformaControl.value = "*";
                plataformaControl.style.display = "inline-block";
            } else {
                plataformaControl.value = "*";
                plataformaControl.style.display = "none";
            }
        }

        let plataforma = plataformaControl.value;

        const url =
            "/forma/snippets/recursos/discos_lista_data.asp" +
            "?folder="      + encodeURIComponent(folder) +
            "&tipo="        + encodeURIComponent(tipo)   +
            "&forma="       + encodeURIComponent(forma)  +
            "&orden="       + encodeURIComponent(orden)  +
            "&amo="         + encodeURIComponent(amo)    +
            "&plataforma="  + encodeURIComponent(plataforma);

        fetch(url)
            .then(r => r.text())
            .then(html => {
                document.getElementById("discos_lista_body").innerHTML = html;
            });
    }    

    function discos_lista_cargarPlataformas(tipo) {
        const combo = document.getElementById("cboPlataforma");
        combo.innerHTML = '<option value="*">- - Plataforma - -</option>';

        const plataformas = discos_plataformas_get();

        plataformas.forEach(p => {
            if (
                (tipo === "JU" && p.juegos === 1) ||
                (tipo === "SO" && p.software === 1)
            ) {

                const opt = document.createElement("option");
                opt.value = p.codigo;
                opt.textContent = p.nombre;

                combo.appendChild(opt);
            }
        });
    }    

    function discos_lista_cargarFormas(tipo) {
        const url =
            "/forma/snippets/recursos/discos_lista_formas.asp" +
            "?tipo=" + encodeURIComponent(tipo);

        fetch(url)
            .then(r => r.text())
            .then(html => {
                const combo = document.getElementById("cboForma");

                combo.innerHTML = '<option value="*">- - Forma - -</option>' + html;
                combo.value = "*";
            });

    }    
</script> 

<% discos_lista_con.close: set discos_lista_con = nothing %>