<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->   
        <%
            thisSystem = "discos"
            thisProcess = "discos.0110"
            SysLockOut

            '
            ' Init()
            '

            dim con, tt, sqlString, cbox, paquete, Objeto
            dim claseLinea, Editor

            Usuario = Request.Cookies("usuario")
            Paquete = Request.QueryString("p")
            Objeto = Request.QueryString("o")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            '
            ' Funciones y procedimientos
            '            

            sub lista_Musica(Usuario, Paquete, Objeto, Editor)
                %>
                    <div class="line label-top">
                        <label class="label tiny2" style="vertical-align: top;">Temas:</label>
                        <div class="label full section">
                            <div class="tabla-wrapper">
                                <table class="tabla tabla-blue">
                                    <thead>
                                        <tr>
                                            <th class="sticky" style="width: 80%; text-align: center;">Canciones</th>
                                            <th class="sticky" style="width: 15%; text-align: center;">Exito</th>
                                            <th class="sticky" style="width:  5%; text-align: center;">&nbsp;</th>
                                        </tr>
                                    </thead>  

                                    <tbody>                              
                                        <%
                                            sqlString = "SELECT Secuencia, Titulo, NumSerieLlave, Exito, Lado " & _
                                                        "FROM discos_Objetos_Detalle " & _
                                                        "WHERE (Usuario = '" & Usuario & "') " & _
                                                        "AND (Paquete = '" & Paquete & "') " & _
                                                        "AND (Objeto = '" & Objeto & "') " & _
                                                    "ORDER BY Titulo;"

                                            set cbox = con.execute(sqlString)
                                                if not (cbox.bof or cbox.eof) then
                                                    Cuantos = 0

                                                    Do
                                                        Cuantos = Cuantos + 1

                                                        %>
                                                            <tr>
                                                                <td>
                                                                    <input class="field" 
                                                                        style="width: 100%; border-color: transparent; background-color: transparent;"
                                                                        name="DM01_FORM_Titulo_<%= Cuantos %>" 
                                                                        id="DM01_FORM_Titulo_<%= Cuantos %>" 
                                                                        type="text" 
                                                                        value="<%= cbox("Titulo") %>" >
                                                                </td>

                                                                <td>
                                                                    <select class="field" 
                                                                            style="width: 100%; border-color: transparent; background-color: transparent;" 
                                                                            name="DM01_FORM_Exito_<%= Cuantos %>" 
                                                                            id="DM01_FORM_Exito_<%= Cuantos %>">
                                                                        <option value="0" <% if cbox("Exito") = 0 then response.write " selected" %>>&nbsp;</option>
                                                                        <option value="1" <% if cbox("Exito") = 1 then response.write " selected" %>>Exito</option>                            
                                                                        <option value="2" <% if cbox("Exito") = 2 then response.write " selected" %>>Aceptable</option>
                                                                    </select>                                                            
                                                                </td>

                                                                <td style="text-align: center;">
                                                                    <button class = "form-btn rojo" 
                                                                            type = "button" 
                                                                            onClick="obj_borrar_det('<%=  cbox("Secuencia") %>', '<%= Editor %>', '<%= paquete %>', '<%= objeto %>')" >
                                                                        <i class="fa fa-trash fa-xl"></i>
                                                                    </button>
                                                                </td>
                                                            </tr>
                                                        <%

                                                        cbox.MoveNext
                                                    Loop Until cbox.eof
                                                end if
                                            cbox.close: set cbox = nothing
                                        %>

                                        <!--
                                            Añadimos un "formulario" para añadir 
                                            más títulos...
                                        -->                                    

                                        <tr>
                                            <td>
                                                <input class="field" 
                                                    style="width: 100%; border-color: transparent; background-color: transparent;" 
                                                    type="text" 
                                                    id="DM01_Titulo" 
                                                    name="DM01_Titulo" placeholder="Nuevo tema" >
                                            </td>

                                            <td>
                                                <select class="field" 
                                                        style="width: 100%; border-color: transparent; background-color: transparent;" 
                                                        name="DM01_Exito" 
                                                        id="DM01_Exito" >
                                                    <option value="0">&nbsp;</option>
                                                    <option value="1">Exito</option>                            
                                                    <option value="2">Aceptable</option>
                                                </select>                                         
                                            </td>

                                            <td style="width: 10%; text-align: center;">
                                                <button class="form-btn verde" 
                                                        type="button" 
                                                        onclick="DM01_NuevaLinea('<%= Paquete %>', '<%= Objeto %>', '<%= Editor %>')">
                                                    <i class="fa fa-save fa-xl"></i>
                                                </button>
                                            </td>
                                        </tr>
                                    </tbody>

                                    <tfoot>
                                        <tr>
                                            <td colspan="3" class="sticky" style="text-align: center;">
                                                <%
                                                    select case Cuantos
                                                        case 0: response.write "El Objeto No Tiene Temas"
                                                        case 1: response.write "El Objeto Tiene Un Temas"
                                                        case else
                                                            response.write "El Objeto Tiene " & Cuantos & " Temas"
                                                    end select
                                                %>
                                            </td>
                                        </tr>
                                    </tfoot>
                                </table> 
                            </div>
                        </div>
                    </div>
                <%     
            end sub

            sub lista_Musica_Multilado(Usuario, Paquete, Objeto, Editor)
                %>
                    <div class="line label-top">
                        <label class="label tiny2" style="vertical-align: top;">Temas:</label>
                        <div class="label full section">
                            <div class="tabla-wrapper">
                                <table class="tabla tabla-blue">
                                    <thead>
                                        <tr>
                                            <th class="sticky" style="width: 10%; text-align: center;">Lado</th>
                                            <th class="sticky" style="width: 70%; text-align: center;">Canciones</th>
                                            <th class="sticky" style="width: 15%; text-align: center;">Exito</th>
                                            <th class="sticky" style="width:  5%; text-align: center;">&nbsp;</th>
                                        </tr>
                                    </thead>  

                                    <tbody>                              
                                        <%
                                            sqlString = "SELECT Secuencia, Titulo, NumSerieLlave, Exito, Lado " & _
                                                        "FROM discos_Objetos_Detalle " & _
                                                        "WHERE (Usuario = '" & Usuario & "') " & _
                                                        "AND (Paquete = '" & Paquete & "') " & _
                                                        "AND (Objeto = '" & Objeto & "') " & _
                                                    "ORDER BY Lado, Titulo;"

                                            set cbox = con.execute(sqlString)
                                                if not (cbox.bof or cbox.eof) then
                                                    Cuantos = 0

                                                    Do
                                                        Cuantos = Cuantos + 1

                                                        %>
                                                            <tr>
                                                                <td>
                                                                    <select class="field" 
                                                                            style="width: 100%; border-color: transparent; background-color: transparent;"
                                                                            name="DM01_FORM_Lado_<%= Cuantos %>" 
                                                                            id="DM01_FORM_Lado_<%= Cuantos %>" >
                                                                        <option value="A" <% if cbox("Lado") = "A" then response.write " selected" %>>A</option>
                                                                        <option value="B" <% if cbox("Lado") = "B" then response.write " selected" %>>B</option>                            
                                                                        <option value="C" <% if cbox("Lado") = "C" then response.write " selected" %>>C</option>
                                                                        <option value="D" <% if cbox("Lado") = "D" then response.write " selected" %>>D</option>
                                                                    </select>                                                            
                                                                </td>

                                                                <td>
                                                                    <input class="field" 
                                                                        style="width: 100%; border-color: transparent; background-color: transparent;"
                                                                        name="DM01_FORM_Titulo_<%= Cuantos %>" 
                                                                        id="DM01_FORM_Titulo_<%= Cuantos %>" 
                                                                        type="text" 
                                                                        value="<%= cbox("Titulo") %>" >
                                                                </td>

                                                                <td>
                                                                    <select class="field" 
                                                                            style="width: 100%; border-color: transparent; background-color: transparent;" 
                                                                            name="DM01_FORM_Exito_<%= Cuantos %>" 
                                                                            id="DM01_FORM_Exito_<%= Cuantos %>">
                                                                        <option value="0" <% if cbox("Exito") = 0 then response.write " selected" %>>&nbsp;</option>
                                                                        <option value="1" <% if cbox("Exito") = 1 then response.write " selected" %>>Exito</option>                            
                                                                        <option value="2" <% if cbox("Exito") = 2 then response.write " selected" %>>Aceptable</option>
                                                                    </select>                                                            
                                                                </td>

                                                                <td style="text-align: center;">
                                                                    <button class = "form-btn rojo" 
                                                                            type = "button" 
                                                                            onClick="obj_borrar_det('<%=  cbox("Secuencia") %>', '<%= Editor %>', '<%= paquete %>', '<%= objeto %>')" >
                                                                        <i class="fa fa-trash fa-xl"></i>
                                                                    </button>
                                                                </td>
                                                            </tr>
                                                        <%

                                                        cbox.MoveNext
                                                    Loop Until cbox.eof
                                                end if
                                            cbox.close: set cbox = nothing
                                        %>

                                        <!--
                                            Añadimos un "formulario" para añadir 
                                            más títulos...
                                        -->                                    

                                        <tr>
                                            <td>
                                                <select class="field" 
                                                        style="width: 100%; border-color: transparent; background-color: transparent;" 
                                                        name="DM02_Lado" 
                                                        id="DM02_Lado" >
                                                    <option value="A" selected>A</option>
                                                    <option value="B">B</option>                            
                                                    <option value="C">C</option>
                                                    <option value="D">D</option>
                                                </select> 
                                            </td>

                                            <td>
                                                <input class="field" 
                                                    style="width: 100%; border-color: transparent; background-color: transparent;" 
                                                    type="text" 
                                                    id="DM02_Titulo" 
                                                    name="DM02_Titulo" 
                                                    placeholder="Nuevo tema" >
                                            </td>

                                            <td>
                                                <select class="field" 
                                                        style="width: 100%; border-color: transparent; background-color: transparent;" 
                                                        name="DM02_Exito" 
                                                        id="DM02_Exito" >
                                                    <option value="0">&nbsp;</option>
                                                    <option value="1">Exito</option>                            
                                                    <option value="2">Aceptable</option>
                                                </select>                                         
                                            </td>

                                            <td style="width: 10%; text-align: center;">
                                                <button class="form-btn verde" 
                                                        type="button" 
                                                        onclick="DM02_NuevaLinea('<%= Paquete %>', '<%= Objeto %>', '<%= Editor %>')">
                                                    <i class="fa fa-save fa-xl"></i>
                                                </button>
                                            </td>
                                        </tr>
                                    </tbody>

                                    <tfoot>
                                        <tr>
                                            <td colspan="4" class="sticky" style="text-align: center;">
                                                <%
                                                    select case Cuantos
                                                        case 0: response.write "El Objeto No Tiene Temas"
                                                        case 1: response.write "El Objeto Tiene Un Temas"
                                                        case else
                                                            response.write "El Objeto Tiene " & Cuantos & " Temas"
                                                    end select
                                                %>
                                            </td>
                                        </tr>
                                    </tfoot>
                                </table> 
                            </div>
                        </div>
                    </div>                
                <%
            end sub

            sub lista_Capitulos(Usuario, Paquete, Objeto, Editor)
                %>
                    <div class="line label-top">
                        <label class="label tiny2" style="vertical-align: top;">Capítulos:</label>
                        <div class="label full section">
                            <div class="tabla-wrapper">
                                <table class="tabla tabla-blue">
                                    <thead>
                                        <tr>
                                            <th class="sticky" style="width: 95%; text-align: center;">Capítulo</th>
                                            <th class="sticky" style="width:  5%; text-align: center;">&nbsp;</th>
                                        </tr>
                                    </thead>  

                                    <tbody>                              
                                        <%
                                            sqlString = "SELECT Secuencia, Titulo, NumSerieLlave, Exito, Lado " & _
                                                          "FROM discos_Objetos_Detalle " & _
                                                         "WHERE (Usuario = '" & Usuario & "') " & _
                                                           "AND (Paquete = '" & Paquete & "') " & _
                                                           "AND (Objeto = '" & Objeto & "') " & _
                                                      "ORDER BY Titulo;"

                                            set cbox = con.execute(sqlString)
                                                if not (cbox.bof or cbox.eof) then
                                                    Cuantos = 0

                                                    Do
                                                        Cuantos = Cuantos + 1

                                                        %>
                                                            <tr>
                                                                <td>
                                                                    <input class="field" 
                                                                           style="width: 100%; border-color: transparent; background-color: transparent;" 
                                                                           name="DM01_FORM_Titulo_<%= Cuantos %>" 
                                                                           id="DM01_FORM_Titulo_<%= Cuantos %>" 
                                                                           type="text"
                                                                           value="<%= cbox("Titulo") %>" >
                                                                </td>

                                                                <td style="text-align: center;">
                                                                    <button class = "form-btn rojo" 
                                                                            type = "button" 
                                                                            onClick="obj_borrar_det('<%=  cbox("Secuencia") %>', '<%= Editor %>', '<%= paquete %>', '<%= objeto %>')" >
                                                                        <i class="fa fa-trash fa-xl"></i>
                                                                    </button>
                                                                </td>
                                                            </tr>
                                                        <%

                                                        cbox.MoveNext
                                                    Loop Until cbox.eof
                                                end if
                                            cbox.close: set cbox = nothing
                                        %>

                                        <!--
                                            Añadimos un "formulario" para añadir 
                                            más títulos...
                                        -->                                    

                                        <tr>
                                            <td>
                                                <input class="field" 
                                                       style="width: 100%; border-color: transparent; background-color: transparent;" 
                                                       type="text" 
                                                       id="DM03_Titulo" 
                                                       name="DM03_Titulo" 
                                                       placeholder = "Nuevo Capítulo" >
                                            </td>

                                            <td style="width: 10%; text-align: center;">
                                                <button class="form-btn verde" 
                                                        type="button" 
                                                        onclick="DM03_NuevaLinea('<%= Paquete %>', '<%= Objeto %>', '<%= Editor %>')" >
                                                    <i class="fa fa-save fa-xl"></i>
                                                </button>
                                            </td>
                                        </tr>
                                    </tbody>

                                    <tfoot>
                                        <tr>
                                            <td colspan="4" class="sticky" style="text-align: center;">
                                                <%
                                                    select case Cuantos
                                                        case 0: response.write "El Objeto No Tiene Capítulos"
                                                        case 1: response.write "El Objeto Tiene Un Capítulo"
                                                        case else
                                                            response.write "El Objeto Tiene " & Cuantos & " Capítulos"
                                                    end select
                                                %>
                                            </td>
                                        </tr>
                                    </tfoot>
                                </table> 
                            </div>
                        </div>
                    </div>                
                <%
            end sub

            sub lista_ItemSeries(Usuario, Paquete, Objeto, Editor)
                %>
                    <div class="line label-top">
                        <label class="label tiny2" style="vertical-align: top;">Llaves:</label>
                        <div class="label full section">
                            <div class="tabla-wrapper">
                                <table class="tabla tabla-blue">
                                    <thead>
                                        <tr>
                                            <th class="sticky" style="width: 50%; text-align: center;">Nombre del Item</th>
                                            <th class="sticky" style="width: 45%; text-align: center;">Número de Serie</th>
                                            <th class="sticky" style="width:  5%; text-align: center;">&nbsp;</th>
                                        </tr>
                                    </thead>  

                                    <tbody>                              
                                        <%
                                            sqlString = "SELECT Secuencia, Titulo, NumSerieLlave, Exito, Lado " & _
                                                          "FROM discos_Objetos_Detalle " & _
                                                         "WHERE (Usuario = '" & Usuario & "') " & _
                                                           "AND (Paquete = '" & Paquete & "') " & _
                                                           "AND (Objeto = '" & Objeto & "') " & _
                                                      "ORDER BY Titulo;"

                                            set cbox = con.execute(sqlString)
                                                if not (cbox.bof or cbox.eof) then
                                                    Cuantos = 0

                                                    Do
                                                        Cuantos = Cuantos + 1

                                                        %>
                                                            <tr>
                                                                <td>
                                                                    <input class="field" 
                                                                           style="width: 100%; border-color: transparent; background-color: transparent;" 
                                                                           name="DM01_FORM_Titulo_<%= Cuantos %>" 
                                                                           id="DM01_FORM_Titulo_<%= Cuantos %>" 
                                                                           type="text" 
                                                                           value="<%= cbox("Titulo") %>" >
                                                                </td>

                                                                <td>
                                                                    <input class="field" 
                                                                           style="width: 100%; border-color: transparent; background-color: transparent;" 
                                                                           name="DM01_FORM_NumSerie_<%= Cuantos %>" 
                                                                           id="DM01_FORM_NumSerie_<%= Cuantos %>" 
                                                                           type="text" value="<%= cbox("NumSerieLlave") %>" >
                                                                </td>

                                                                <td style="text-align: center;">
                                                                    <button class = "form-btn rojo" 
                                                                            type = "button" 
                                                                            onClick="obj_borrar_det('<%= cbox("Secuencia") %>', '<%= Editor %>', '<%= paquete %>', '<%= objeto %>')" >
                                                                        <i class="fa fa-trash fa-xl"></i>
                                                                    </button>
                                                                </td>
                                                            </tr>
                                                        <%

                                                        cbox.MoveNext
                                                    Loop Until cbox.eof
                                                end if
                                            cbox.close: set cbox = nothing
                                        %>

                                        <!--
                                            Añadimos un "formulario" para añadir 
                                            más títulos...
                                        -->                                    

                                        <tr>
                                            <td>
                                                <input class="field" 
                                                       style="width: 100%; border-color: transparent; background-color: transparent;" 
                                                       type="text" 
                                                       id="DM04_Titulo" 
                                                       name="DM04_Titulo" 
                                                       placeholder = "Nuevo Item" >
                                            </td>

                                            <td>
                                                <input class="field" 
                                                       style="width: 100%; border-color: transparent; background-color: transparent;" 
                                                       type="text" 
                                                       id="DM04_NumSerieLlave" 
                                                       name="DM04_NumSerieLlave" 
                                                       placeholder = "Nuevo Número de Serie" >
                                            </td>

                                            <td style="width: 10%; text-align: center;">
                                                <button class="form-btn verde" 
                                                        type="button" 
                                                        onclick="DM04_NuevaLinea('<%= Paquete %>', '<%= Objeto %>', '<%= Editor %>')" >
                                                    <i class="fa fa-save fa-xl"></i>
                                                </button>
                                            </td>
                                        </tr>
                                    </tbody>

                                    <tfoot>
                                        <tr>
                                            <td colspan="3" class="sticky" style="text-align: center;">
                                                <%
                                                    select case Cuantos
                                                        case 0: response.write "El Objeto No Tiene Items"
                                                        case 1: response.write "El Objeto Tiene Un Item"
                                                        case else
                                                            response.write "El Objeto Tiene " & Cuantos & " Items"
                                                    end select
                                                %>
                                            </td>
                                        </tr>
                                    </tfoot>
                                </table> 
                            </div>
                        </div>
                    </div>                
                <%
            end sub

            sub lista_Capitulos2(Usuario, Paquete, Objeto, Editor)
                %>
                    <div class="line label-top">
                        <label class="label tiny2" style="vertical-align: top;">Capítulos:</label>
                        <div class="label full section">
                            <div class="tabla-wrapper">
                                <table class="tabla tabla-blue">
                                    <thead>
                                        <tr>
                                            <th class="sticky" style="width: 95%; text-align: center;">Capítulos</th>
                                            <th class="sticky" style="width:  5%; text-align: center;">&nbsp;</th>
                                        </tr>
                                    </thead>  

                                    <tbody>                              
                                        <%
                                            sqlString = "SELECT Secuencia, Titulo, NumSerieLlave, Exito, Lado " & _
                                                          "FROM discos_Objetos_Detalle " & _
                                                         "WHERE (Usuario = '" & Usuario & "') " & _
                                                           "AND (Paquete = '" & Paquete & "') " & _
                                                           "AND (Objeto = '" & Objeto & "') " & _
                                                      "ORDER BY Titulo;"

                                            set cbox = con.execute(sqlString)
                                                if not (cbox.bof or cbox.eof) then
                                                    Cuantos = 0

                                                    Do
                                                        Cuantos = Cuantos + 1

                                                        %>
                                                            <tr>
                                                                <td>
                                                                    <input class="field" 
                                                                           style="width: 100%; border-color: transparent; background-color: transparent;" 
                                                                           name="DM01_FORM_Titulo_<%= Cuantos %>" 
                                                                           id="DM01_FORM_Titulo_<%= Cuantos %>" 
                                                                           type="text" 
                                                                           value="<%= cbox("Titulo") %>" >
                                                                </td>

                                                                <td style="text-align: center;">
                                                                    <button class = "form-btn rojo" 
                                                                            type = "button" 
                                                                            onClick="obj_borrar_det('<%= cbox("Secuencia") %>', '<%= Editor %>', '<%= paquete %>', '<%= objeto %>')" >
                                                                        <i class="fa fa-trash fa-xl"></i>
                                                                    </button>
                                                                </td>
                                                            </tr>
                                                        <%

                                                        cbox.MoveNext
                                                    Loop Until cbox.eof
                                                end if
                                            cbox.close: set cbox = nothing
                                        %>

                                        <!--
                                            Añadimos un "formulario" para añadir 
                                            más títulos...
                                        -->                                    

                                        <tr>
                                            <td>
                                                <input class="field" 
                                                       style="width: 100%; border-color: transparent; background-color: transparent;" 
                                                       type="text" 
                                                       id="DM05_Titulo" 
                                                       name="DM05_Titulo" 
                                                       placeholder = "Nuevo Capítulo" >
                                            </td>

                                            <td style="width: 10%; text-align: center;">
                                                <button class="form-btn verde" 
                                                        type="button" 
                                                        onclick="DM05_NuevaLinea('<%= Paquete %>', '<%= Objeto %>', '<%= Editor %>')" >
                                                    <i class="fa fa-save fa-xl"></i>
                                                </button>
                                            </td>
                                        </tr>
                                    </tbody>

                                    <tfoot>
                                        <tr>
                                            <td colspan="2" class="sticky" style="text-align: center;">
                                                <%
                                                    select case Cuantos
                                                        case 0: response.write "El Objeto No Tiene Capítulos"
                                                        case 1: response.write "El Objeto Tiene Un Capítulo"
                                                        case else
                                                            response.write "El Objeto Tiene " & Cuantos & " Capítulos"
                                                    end select
                                                %>
                                            </td>
                                        </tr>
                                    </tfoot>
                                </table> 
                            </div>
                        </div>
                    </div>                
                <%
            end sub

            sub ListaProtagonistas(Usuario, Paquete, Objeto, Editor)
                %>
                    <div class="line label-top">
                        <label class="label tiny2" style="vertical-align: top;">Protagonistas:</label>
                        <div class="label full section">
                            <div class="tabla-wrapper">
                                <table class="tabla tabla-violet">
                                    <thead>
                                        <tr>
                                            <th class="sticky" style="width: 95%; text-align: center;">Actor / Actriz</th>
                                            <th class="sticky" style="width:  5%; text-align: center;">&nbsp;</th>
                                        </tr>
                                    </thead>  

                                    <tbody>                              
                                        <%
                                            sqlString = "SELECT Secuencia, Protagonista " & _
                                                          "FROM discos_Objetos_Protagonistas " & _
                                                         "WHERE (Usuario = '" & Usuario & "') " & _
                                                           "AND (Paquete = '" & Paquete & "') " & _
                                                           "AND (Objeto = '" & Objeto & "') " & _
                                                      "ORDER BY Protagonista;"

                                            set cbox = con.execute(sqlString)
                                                if not (cbox.bof or cbox.eof) then
                                                    Cuantos = 0

                                                    Do
                                                        Cuantos = Cuantos + 1

                                                        %>
                                                            <tr>
                                                                <td>
                                                                    <input class="field" 
                                                                           style="width: 100%; border-color: transparent; background-color: transparent;" 
                                                                           name="DM01_FORM_Protagonista_<%= Cuantos %>" 
                                                                           id="DM01_FORM_Protagonista_<%= Cuantos %>" 
                                                                           type="text" 
                                                                           value="<%= cbox("Protagonista") %>" >
                                                                </td>

                                                                <td style="text-align: center;">
                                                                    <button class = "form-btn rojo" 
                                                                            type = "button" 
                                                                            onClick="obj_borrar_prot('<%=  cbox("Secuencia") %>', '<%= Editor %>', '<%= paquete %>', '<%= objeto %>')" >
                                                                        <i class="fa fa-trash fa-xl"></i>
                                                                    </button>
                                                                </td>
                                                            </tr>
                                                        <%

                                                        cbox.MoveNext
                                                    Loop Until cbox.eof
                                                end if
                                            cbox.close: set cbox = nothing
                                        %>

                                        <!--
                                            Añadimos un "formulario" para añadir 
                                            más títulos...
                                        -->                                    

                                        <tr>
                                            <td>
                                                <input class="field" 
                                                       style="width: 100%; border-color: transparent; background-color: transparent;" 
                                                       type="text" 
                                                       id="DM06_Protagonista" 
                                                       name="DM06_Protagonista" 
                                                       placeholder = "Nuevo Actor o Actriz" >
                                            </td>

                                            <td style="width: 10%; text-align: center;">
                                                <button class="form-btn verde" 
                                                        type="button"
                                                        onclick="DM06_NuevaLinea('<%= Paquete %>', '<%= Objeto %>', '<%= Editor %>')" >
                                                    <i class="fa fa-save fa-xl"></i>
                                                </button>
                                            </td>
                                        </tr>
                                    </tbody>

                                    <tfoot>
                                        <tr>
                                            <td colspan="2" class="sticky" style="text-align: center;">
                                                <%
                                                    select case Cuantos
                                                        case 0: response.write "El Objeto No Tiene Protagonistas"
                                                        case 1: response.write "El Objeto Tiene Un Protagonista"
                                                        case else
                                                            response.write "El Objeto Tiene " & Cuantos & " Protagonistas"
                                                    end select
                                                %>
                                            </td>
                                        </tr>
                                    </tfoot>
                                </table> 
                            </div>
                        </div>
                    </div>                
                <%
            end sub

            sub ListaIdiomasPelicula(Usuario, Paquete, Objeto, Editor)
                %>
                    <div class="line label-top">
                        <label class="label tiny2" style="vertical-align: top;">Idiomas:</label>
                        <div class="label full section">
                            <div class="tabla-wrapper">
                                <table class="tabla tabla-red">
                                    <thead>
                                        <tr>
                                            <th class="sticky" style="width: 50%; text-align: center;">Idioma</th>
                                            <th class="sticky" style="width: 20%; text-align: center;">Audio</th>
                                            <th class="sticky" style="width: 20%; text-align: center;">Subtítulos</th>
                                            <th class="sticky" style="width: 10%; text-align: center;">&nbsp;</th>                         
                                        </tr>
                                    </thead>  

                                    <tbody>                              
                                        <%
                                            sqlString = "SELECT oi.Idioma AS CodigoIdioma, oi.Audio, oi.Subtitulos, i.Nombre AS Idioma, oi.Secuencia " & _
                                                          "FROM discos_Objetos_Idiomas AS oi " & _
                                                    "INNER JOIN discos_Idiomas AS i " & _
                                                            "ON oi.Usuario = i.Usuario " & _
                                                           "AND oi.Idioma = i.Codigo " & _
                                                         "WHERE (oi.Usuario = '" & Usuario & "') " & _
                                                           "AND (oi.Paquete = '" & Paquete & "') " & _
                                                           "AND (oi.Objeto = '" & Objeto & "') " & _
                                                      "ORDER BY Idioma;"

                                            set cbox = con.execute(sqlString)
                                                if not (cbox.bof or cbox.eof) then
                                                    Cuantos = 0

                                                    Do
                                                        Cuantos = Cuantos + 1

                                                        %>
                                                            <tr>
                                                                <td><%= cbox("Idioma") %></td>

                                                                <td>
                                                                    <select class="field" 
                                                                            style="width: 100%; border-color: transparent; background-color: transparent;"
                                                                            name="DM07_Audio_<%= Cuantos %>" 
                                                                            id="DM07_Audio_<%= Cuantos %>" required >
                                                                        <option value="1" <% if cbox("Audio") = "1" then response.write " selected" %>>Sí</option>
                                                                        <option value="0" <% if cbox("Audio") = "0" then response.write " selected" %>>&nbsp;</option>
                                                                    </select>                                                                
                                                                </td>

                                                                <td>
                                                                    <select class="field" 
                                                                            style="width: 100%; border-color: transparent; background-color: transparent;"
                                                                            name="DM07_SubTitulo_<%= Cuantos %>" 
                                                                            id="DM07_SubTitulo_<%= Cuantos %>" required >
                                                                        <option value="1" <% if cbox("SubTitulos") = "1" then response.write " selected" %>>Sí</option>
                                                                        <option value="0" <% if cbox("SubTitulos") = "0" then response.write " selected" %>>&nbsp;</option>
                                                                    </select>                                                                
                                                                </td>

                                                                <td style="text-align: center;">
                                                                    <button class = "form-btn rojo" 
                                                                            type = "button" 
                                                                            onClick="obj_borrar_idioma('<%=  cbox("Secuencia") %>', '<%= Editor %>', '<%= paquete %>', '<%= objeto %>')" >
                                                                        <i class="fa fa-trash fa-xl"></i>
                                                                    </button>
                                                                </td>
                                                            </tr>
                                                        <%

                                                        cbox.MoveNext
                                                    Loop Until cbox.eof
                                                end if
                                            cbox.close: set cbox = nothing
                                        %>

                                        <!--
                                            Añadimos un "formulario" para añadir 
                                            más títulos...
                                        -->                                    

                                        <tr>
                                            <td>
                                                <select class="field" 
                                                        style="width: 100%; border-color: transparent; background-color: transparent;"
                                                        name='DM07_Idiomaa' id='DM07_Idiomaa' required >

                                                        <%
                                                            sqlString = "SELECT Codigo, Nombre " & _
                                                                        "FROM discos_Idiomas " & _
                                                                        "WHERE (Usuario = '" & Usuario & "')" & _
                                                                    "ORDER BY Nombre;"

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
                                            </td>

                                            <td>
                                                <select class="field" 
                                                        style="width: 100%; border-color: transparent; background-color: transparent;"
                                                        name="DM07_Audio" 
                                                        id="DM07_Audio" required >
                                                    <option value="1">Sí</option>
                                                    <option value="0">&nbsp;</option>
                                                </select>
                                            </td>

                                            <td>
                                                <select class="field" 
                                                        style="width: 100%; border-color: transparent; background-color: transparent;"
                                                        name="DM07_SubTitulos" 
                                                        id="DM07_SubTitulos" required >
                                                    <option value="1">Sí</option>
                                                    <option value="0">&nbsp;</option>
                                                </select>
                                            </td>                                            

                                            <td style="width: 10%; text-align: center;">
                                                <button class="form-btn verde" 
                                                        type="button"
                                                        onclick="DM07_NuevaLinea('<%= Paquete %>', '<%= Objeto %>', '<%= Editor %>')" >
                                                    <i class="fa fa-save fa-xl"></i>
                                                </button>
                                            </td>
                                        </tr>
                                    </tbody>

                                    <tfoot>
                                        <tr>
                                            <td colspan="4" class="sticky" style="text-align: center;">
                                                <%
                                                    select case Cuantos
                                                        case 0: response.write "El Objeto No Tiene Idiomas"
                                                        case 1: response.write "El Objeto Tiene Un Idioma"
                                                        case else
                                                            response.write "El Objeto Tiene " & Cuantos & " Idiomas"
                                                    end select
                                                %>
                                            </td>
                                        </tr>
                                    </tfoot>
                                </table> 
                            </div>
                        </div>
                    </div>                
                <%
            end sub
        %>  

        <style>
            .line-group .line {
                margin-bottom: 12px;
            }

            .line-group .line:last-child {
                margin-bottom: 0;
            }      

            .img-limitada {
                display: block;
                margin: 0 auto;
                width: 100%;
                height: auto;
                max-height: 500px;
                min-height: 50px;
                object-fit: contain;
            }

            .campo {
                padding: 0.3rem 0.4rem;
                border: 1px solid #ccc;
                border-radius: 0.3rem;
                font-family: 'Ruda', sans-serif;
                font-size: 1rem;
                color: rgb(25, 25, 25);
                box-sizing: border-box;
                resize: vertical;
            }            
            
            .label.tiny2        { width: 120px ; }
            .field.año          { width: 75px ; }
            .field.description  { width: 550px; }  
            .field.normal_plus  { width: 400px; }
        </style>
    </head>

    <body plantilla="normal" reserva="160">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            sqlString = "SELECT o.AEdicion, o.Titulo, o.TituloOriginal, o.IdiomaMusica, o.InDirAu, o.DuraPag, " & _
                              " o.Forma, o.Clasificacion, o.Tipo, o.FormatoPantalla, o.Recuento, o.Es3D, o.PlatOS, " & _
                              " o.CopiaDigital, o.Descripcion, o.Editor, o.Visible, f.Multilados, f.Icono_Forma " & _
                          "FROM dbo.discos_Objetos AS o " & _
                    "INNER JOIN dbo.discos_Formas AS f " & _
                            "ON o.Usuario = f.Usuario " & _
                           "AND o.Forma = f.Forma " & _
                         "WHERE (o.Usuario = '" & Usuario & "') " & _
                           "AND (o.Paquete = '" & Paquete & "') " & _
                           "AND (o.Objeto = '" & Objeto & "');"

            set tt = con.execute(sqlString)        
        %>

        <br />
        
        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 70%; text-align: left; font-size: 18px; color: rgb(50, 50, 50);">
                <%
                    img_Forma = lcase(Request.Cookies("usuPath")) & "/discos/" & tt("Icono_Forma")
                    img_Editor = lcase(Request.Cookies("usuPath")) & "/discos/ob_" & tt("Editor") & ".jpg"
                    img_3d = lcase(Request.Cookies("usuPath")) & "/discos/icono_3D.gif"

                    'response.write "<img src='" & img_Editor & "' height='60' >&nbsp;"

                    if tt("Es3D") = 1 then
                        'response.write "<img src='" & img_3D & "'  height='60'>&nbsp;&nbsp;"
                    end if

                    'response.write "<img src='" & img_Forma & "' height='60' >&nbsp;&nbsp;"
                %>

                <span style="font-size: 20px;">
                    <%= "(" & Objeto & ") " & tt("Titulo") %>
                </span>
            </div>
            
            <div style="flex: 0 0 30%; text-align: right;">
                <button type="button" class="form-btn verde normal" onclick="enviar()" >
                    Grabar
                </button>    

                <button type="button" class="form-btn azul normal" onclick="irA('lista.asp')" >
                    Cancelar
                </button>                        
            </div>
        </div>        

        <form name="form_transaccion" id="form_transaccion" method="post" action="grabar_objeto.asp">
            <div class="main main-scroll">
                <!-- Parte 1: Generales del Paquete -->
                    <div class="no-ver">
                        <input type="text" id="Paquete" name="Paquete" value="<%= Paquete %>" >
                        <input type="text" id="Objeto"  name="Objeto"  value="<%= Objeto %>" >
                        <input type="text" id="Editor"  name="Editor"  value="<%= tt("Editor") %>" >
                    </div>

                    <table class="tabla-transparente" style="width:100%; table-layout: fixed;">
                        <tr>
                            <td style="width:65%;">
                                <div class="line-group">
                                    <div class="line">
                                        <label class="label tiny2" onclick="Filtro('Amo = <%= tt("AEdicion") %>', 'A&ntilde;o = <%= tt("AEdicion") %>')" >Año Edición</label>
                                        <input class="field año" type="text" id="AEdicion" name="AEdicion" value="<%= tt("AEdicion") %>" placeholder="9999" required>
                                    </div>

                                    <div class="line">
                                        <label class="label tiny2">Título</label>
                                        <input class="field description" 
                                            type="text" id="Titulo" name="Titulo" 
                                            value="<%= tt("Titulo") %>" 
                                            placeholder="Título del Paquete" required >
                                    </div>

                                    <% if tt("Editor") = "PE" then %>
                                        <div class="line">
                                            <label class="label tiny2">Título Original</label>
                                            <input class="field description" 
                                                type="text" id="TituloOriginal" name="TituloOriginal" 
                                                value="<%= tt("TituloOriginal") %>" 
                                                placeholder="Título Original del Paquete" required >
                                        </div>
                                    <% end if %>     

                                    <% if tt("Editor") <> "HW" then %>
                                        <div class="line">
                                            <%
                                                ida_label = ""

                                                select case tt("Editor")
                                                    case "DM", "VM": ida_label = "Intérprete"
                                                    case "PE", "JU": ida_label = "Director"
                                                    case "LI", "SO": ida_label = "Autor"
                                                end select

                                                ida_filtro = "CHARINDEX(¿!" & tt("InDirAu") & "¿!, ListaInDirAu) > 0"
                                                ida_titulo = ida_label & " = " & tt("InDirAu")
                                            %>  

                                            <label class="label tiny2" onclick="Filtro('<%= ida_filtro %>', '<%= ida_titulo %>')"><%= ida_label %></label>

                                            <input class="field description" type="text" id="InDirAu" name="InDirAu" 
                                                value="<%= tt("InDirAu") %>" placeholder="<%= ida_label %> del Paquete" required>
                                        </div>
                                    <% end if %>

                                    <div class="line">
                                        <%
                                            ida_filtro = "CHARINDEX(¿!" & tt("Tipo") & "¿!, ListaGeneros) > 0"
                                            ida_titulo = "Género = " & tt("Tipo")
                                        %>

                                        <label class="label tiny2" onclick="Filtro('<%= ida_filtro %>', '<%= ida_titulo %>')" >
                                            <% 
                                                if tt("Editor") <> "HW" then 
                                                    response.write "Género"
                                                else
                                                    response.write "Equipo"
                                                end if
                                            %>                                    
                                        </label>

                                        <%
                                            sqlString = "SELECT Codigo, Nombre " & _
                                                        "FROM discos_Tipos " & _
                                                        "WHERE (Usuario = '" & usuario & "') " & _
                                                        "AND (codigo <> '00000000') " 
                                            
                                            select case tt("Editor")
                                                case "DM", "VM": sqlString = sqlString & "AND (Musica = 1) "
                                                case "PE": sqlString = sqlString & "AND (Video = 1) "
                                                case "JU": sqlString = sqlString & "AND (Juegos = 1) "
                                                case "SO": sqlString = sqlString & "AND (Software = 1) "
                                                case "LI": sqlString = sqlString & "AND (Libros = 1) "
                                                case "HW": sqlString = sqlString & "AND (Hardware = 1) "
                                            end select

                                            sqlString = sqlString & "ORDER BY Nombre;"

                                            set cbox = con.execute(sqlString)
                                                if not (cbox.bof or cbox.eof) then
                                                    response.write "<select class='field description' name='Tipo' id='Tipo' required >" 
                                                        Do
                                                            response.write "<option value='" & cbox("Codigo") & "' "
                                                                if tt("Tipo") = cbox("Codigo") then response.write " selected" 
                                                            response.write ">" & cbox("Nombre") & "</option>"

                                                            cbox.MoveNext
                                                        Loop Until cbox.eof
                                                    response.write "</select>"
                                                end if
                                            cbox.close: set cbox = nothing
                                        %>                                      
                                    </div>

                                    <div class="line">
                                        <%
                                            ida_filtro = "CHARINDEX(¿!" & tt("Forma") & "¿!, ListaFormas) > 0"
                                            ida_titulo = "Forma = " & tt("Forma")
                                        %>

                                        <label class="label tiny2" onclick="Filtro('<%= ida_filtro %>', '<%= ida_titulo %>')" >Forma</label>

                                        <%
                                            sqlString = "SELECT Forma, Nombre " & _
                                                        "FROM discos_Formas " & _ 
                                                        "WHERE (Usuario = '" & Usuario & "') " 

                                            select case tt("Editor")
                                                case "DM": sqlString = sqlString & "AND (Musica = 1) "
                                                case "VM": sqlString = sqlString & "AND (Video = 1) "
                                                case "PE": sqlString = sqlString & "AND (Video = 1) "
                                                case "JU": sqlString = sqlString & "AND (Juegos = 1) "
                                                case "SO": sqlString = sqlString & "AND (Software = 1) "
                                                case "LI": sqlString = sqlString & "AND (Libros = 1) "
                                                case "HW": sqlString = sqlString & "AND (HArdware = 1) "
                                            end select

                                            sqlString = sqlString & "ORDER BY Nombre;"                                    

                                            set cbox = con.execute(sqlString)
                                                if not (cbox.bof or cbox.eof) then
                                                    response.write "<select class='field description' name='Forma' id='Forma' required >" 
                                                        Do
                                                            response.write "<option value='" & cbox("Forma") & "' "
                                                                if cbox("Forma") = tt("Forma") then response.write " selected" 
                                                            response.write ">" & cbox("Nombre") & "</option>"

                                                            cbox.MoveNext
                                                        Loop Until cbox.eof
                                                    response.write "</select>"
                                                end if
                                            cbox.close: set cbox = nothing
                                        %>                                    
                                    </div>

                                    <% if tt("Editor") = "DM" OR tt("Editor") = "VM" then %>
                                        <div class="line">
                                            <label class="label tiny2">Tipo</label>

                                            <select class="field normal" name="Recuento" id="Recuento" required >
                                                <% if tt("Editor") = "DM" then %>
                                                    <option value="1" <% if tt("Recuento") = "1" then response.write " selected" %>>Álbum</option>                            
                                                <% end if %>
                                            
                                                <option value="2" <% if tt("Recuento") = "2" then response.write " selected" %>>Recuento de Éxitos</option>
                                                <option value="3" <% if tt("Recuento") = "3" then response.write " selected" %>>Concierto en Vivo</option>             
                                            </select>     
                                        </div>
                                    <% end if %>

                                    <div class="line">
                                        <label class="label tiny2">Descripción</label>
                                        <textarea class="field description" 
                                                type="text" id="Descripcion" name="Descripcion" 
                                                rows = 5 cols = 80 ><%= tt("Descripcion") %></textarea>
                                    </div>                                

                                    <% if tt("Editor") = "PE" or tt("Editor") = "VM" or tt("Editor") = "LI" then %>
                                        <div class="line">
                                            <label class="label tiny2">
                                                <%
                                                    Select Case tt("Editor")
                                                        Case "PE", "VM": response.write "Duración"
                                                        Case "LI": response.write "Páginas"
                                                    End Select
                                                %>                         
                                            </label>

                                            <input class="field tiny" type="text" name="DuraPag" id="DuraPag" value="<%= tt("DuraPag") %>">
                                        </div>
                                    <% end if %>

                                    <% 
                                        if tt("Editor") = "PE" or tt("Editor") = "VM" then 
                                            ida_filtro = "Medio3D = " & tt("Es3D")
                                            
                                            select case tt("Es3D")
                                                Case 1: ida_titulo = "Medios En 3-D"
                                                Case 0: ida_titulo = "Medios que no son en 3-D"
                                            end select
                                     %>
                                        <div class="line">
                                            <label class="label tiny2" onclick="Filtro('<%= ida_filtro %>', '<%= ida_titulo %>')">Es 3D</label>

                                            <select class="field tiny" name="Es3D" id="Es3D" required >
                                                <option value="0" <% if tt("Es3D") = "0" then response.write " selected" %>>No</option>                            
                                                <option value="1" <% if tt("Es3D") = "1" then response.write " selected" %>>Es 3-D</option>             
                                            </select>                                         
                                        </div>
                                    <% end if %>   

                                    <% if tt("Editor") = "PE" or tt("Editor") = "VM" then %>
                                        <div class="line">
                                            <label class="label tiny2">Pantalla</label>

                                            <%
                                                sqlString = "SELECT Codigo, Nombre FROM discos_FormatosPantalla " & _
                                                            "WHERE Usuario = '" &  Usuario & "' AND Codigo <> '00000000' " & _
                                                        "ORDER BY Nombre;"

                                                set cbox = con.execute(sqlString)
                                                    if not (cbox.bof or cbox.eof) then
                                                        response.write "<select class='field normal_plus' name='FormatoPantalla' id='FormatoPantalla' required >" 
                                                            Do
                                                                response.write "<option value='" & cbox("Codigo") & "' "
                                                                    if cbox("Codigo") = tt("FormatoPantalla") then response.write " selected" 
                                                                response.write ">" & cbox("Nombre") & "</option>"

                                                                cbox.MoveNext
                                                            Loop Until cbox.eof
                                                        response.write "</select>"
                                                    end if
                                                cbox.close: set cbox = nothing
                                            %>                                      
                                        </div>
                                    <% end if %>   

                                    <% if tt("Editor") = "PE" or tt("Editor") = "VM" OR tt("Editor") = "JU" then %>
                                        <div class="line">
                                            <label class="label tiny2">Clasificación</label>

                                            <%
                                                sqlString = "SELECT Codigo, Nombre FROM discos_Clasificaciones " & _
                                                            "WHERE Usuario = '" & Usuario & "' AND Codigo <> '-' " & _
                                                        "ORDER BY Nombre;"

                                                set cbox = con.execute(sqlString)
                                                    if not (cbox.bof or cbox.eof) then
                                                        response.write "<select class='field normal_plus' name='Clasificacion' id='Clasificacion' required >" 
                                                            Do
                                                                response.write "<option value='" & cbox("Codigo") & "' "
                                                                    if cbox("Codigo") = tt("Clasificacion") then response.write " selected" 
                                                                response.write ">" & cbox("Nombre") & "</option>"

                                                                cbox.MoveNext
                                                            Loop Until cbox.eof
                                                        response.write "</select>"
                                                    end if
                                                cbox.close: set cbox = nothing
                                            %>                                   
                                        </div>
                                    <% end if %>    

                                    <% if tt("Editor") = "SO" OR tt("Editor") = "JU" then %>
                                        <div class="line">
                                            <label class="label tiny2">Plataforma</label>

                                            <%
                                                sqlString = "SELECT Codigo, Nombre FROM discos_Plataformas " & _
                                                            "WHERE (Usuario = '" & Usuario & "') " & _
                                                            "AND (Codigo <> '00000000') " 

                                                select case tt("Editor")
                                                    case "JU": sqlString = sqlString & " AND (Juegos = 1) " 
                                                    case "SO": sqlString = sqlString & " AND (Software = 1) " 
                                                end select

                                                sqlString = sqlString & "ORDER BY Nombre;"

                                                set cbox = con.execute(sqlString)
                                                    if not (cbox.bof or cbox.eof) then
                                                        response.write "<select class='field normal_plus' name='PlatOS' id='PlatOS' required >" 
                                                            Do
                                                                response.write "<option value='" & cbox("Codigo") & "' "
                                                                    if cbox("Codigo") = tt("PlatOS") then response.write " selected" 
                                                                response.write ">" & cbox("Nombre") & "</option>"

                                                                cbox.MoveNext
                                                            Loop Until cbox.eof
                                                        response.write "</select>"
                                                    end if
                                                cbox.close: set cbox = nothing
                                            %>                                 
                                        </div>
                                    <% end if %>          

                                    <% if tt("Editor") = "DM" OR tt("Editor") = "VM" then %>
                                        <div class="line">
                                            <label class="label tiny2">Idioma</label>

                                            <%
                                                sqlString = "SELECT Codigo, Nombre " & _
                                                            "FROM discos_Idiomas " & _
                                                            "WHERE (Usuario = '" & Usuario & "')" & _
                                                        "ORDER BY Nombre;"

                                                set cbox = con.execute(sqlString)
                                                    if not (cbox.bof or cbox.eof) then
                                                        response.write "<select class='field normal_plus' name='IdiomaMusica' id='IdiomaMusica' required >" 
                                                            Do
                                                                response.write "<option value='" & cbox("Codigo") & "' "
                                                                    if tt("IdiomaMusica") = cbox("Codigo") then response.write " selected" 
                                                                response.write ">" & cbox("Nombre") & "</option>"

                                                                cbox.MoveNext
                                                            Loop Until cbox.eof
                                                        response.write "</select>"
                                                    end if
                                                cbox.close: set cbox = nothing
                                            %>                                 
                                        </div>
                                    <% end if %>     

                                    <% if tt("Editor") <> "HW" then %>
                                        <div class="line">
                                            <label class="label tiny2">Copia Digital</label>

                                            <%
                                                sqlString = "SELECT Codigo, Nombre " & _
                                                            "FROM discos_Tiendas " & _
                                                            "WHERE (Usuario = '" & Usuario & "') " & _
                                                            "AND (MediosDigitales = 1) " 

                                                select case tt("Editor")
                                                    case "DM", "VM": sqlString = sqlString & "AND (Musica = 1) "
                                                    case "PE": sqlString = sqlString & "AND (Video = 1) "
                                                    case "JU": sqlString = sqlString & "AND (Juegos = 1) "
                                                    case "SO": sqlString = sqlString & "AND (Software = 1) "
                                                    case "LI": sqlString = sqlString & "AND (Libros = 1) "
                                                end select

                                                sqlString = sqlString & "ORDER BY Nombre;"                                                           

                                                set cbox = con.execute(sqlString)
                                                    if not (cbox.bof or cbox.eof) then
                                                        response.write "<select class='field normal_plus' name='CopiaDigital' id='CopiaDigital' required >" 
                                                            response.write "<option value='' " 
                                                                if tt("CopiaDigital") = "" then 
                                                                    response.write " selected" 
                                                                end if
                                                            response.write ">&nbsp;</option>"

                                                            Do
                                                                response.write "<option value='" & cbox("Codigo") & "' "
                                                                    if tt("CopiaDigital") = cbox("Codigo") then response.write " selected" 
                                                                response.write ">" & cbox("Nombre") & "</option>"

                                                                cbox.MoveNext
                                                            Loop Until cbox.eof
                                                        response.write "</select>"
                                                    end if
                                                cbox.close: set cbox = nothing
                                            %>                                 
                                        </div>
                                    <% end if %>
                                </div>                          
                            </td>

                            <td style="width:35%; vertical-align: top; text-align: center;">
                                <%
                                    if (nuevo = 1) then
                                        response.write "<img src='/core/imagenes/misc/foto.jpg' alt='Portada del Objeto'>"
                                    else
                                        fotoObjeto = request.Cookies("usuPath") & "/medios/" & Objeto & ".jpg"
                                        fotoPaquete = request.Cookies("usuPath") & "/medios/" & Paquete & ".jpg"

                                        %>
                                        <img class="img-limitada" 
                                                name="portada" id="portada" src="<%= fotoObjeto %>" 
                                                onerror="this.src='<%= fotoPaquete %>'" 
                                                onclick="NuevaFoto('<%= Paquete %>', '<%= Objeto %>')" >
                                        <%
                                    end if
                                %>                            
                            </td>
                        </tr>
                    </table>
                <!-- Fin - Parte 1: Generales del Paquete -->

                <br /><br />

                <!-- Parte 2: Secciones del Objeto -->                    
                    <%
                        select Case tt("Editor")
                            case "DM"
                                if tt("Multilados") = 1 then
                                    lista_Musica_Multilado Usuario, Paquete, Objeto, "DM"
                                else
                                    lista_Musica  Usuario, Paquete, Objeto, "DM"
                                end if

                            case "VM"
                                lista_Musica Usuario, Paquete, Objeto, "VM"

                            case "PE"
                                lista_Capitulos2 Usuario, Paquete, Objeto, "PE"

                                response.write "<br />"

                                ListaProtagonistas Usuario, Paquete, Objeto, "PE"

                                response.write "<br />"

                                ListaIdiomasPelicula Usuario, Paquete, Objeto, "PE" 

                            Case "LI"
                                lista_Capitulos Usuario, Paquete, Objeto, "LI"

                            case "JU"
                                lista_ItemSeries Usuario, Paquete, Objeto, "JU"

                            case "SO"
                                lista_ItemSeries Usuario, Paquete, Objeto, "SO"
                        end select
                    %> 
                <!-- Fin - Parte 2: Secciones del Objeto -->
            </div>
        </form>    

        <br /><br />

        <script type="text/javascript">
            function irA(vinculo) {
                window.location.href = vinculo;
            }

            function enviar() {
                document.getElementById("form_transaccion").submit();
            }

            function NuevaMetaData(paquete) {
                var meta = document.getElementById("NuevaMetaData").value;
                var vinculo = "metadata_grabar.asp?p=" + paquete + "&m=" + meta;

                window.location.href = vinculo;
            }

            function BorrarMetaData(paquete, metadata) {
                var confirmacion = confirm("Desea borrar la MetaData " + metadata + "?");
                var vinculo = "metadata_borrar.asp?p=" + paquete + "&m=" + metadata;

                if (confirmacion) {     
                window.location.href = vinculo;
                };
            }     

            function NuevaFoto(paquete, objeto) {
                var vinculo = "editar_objeto_foto.asp?p=" + paquete + "&o=" + objeto;
                window.location.href = vinculo;          
            }  

            function DM01_NuevaLinea(paquete, objeto, editor) {
                var tit2 = document.getElementById("DM01_Titulo").value;
                var tit = tit2.replace(/&/g, "yy");
                var exi = document.getElementById("DM01_Exito").value;

                if (tit != "") {
                    var vinculo = "objeto_grabar_detalle.asp?op=1&p=" + paquete + "&o=" + objeto + "&ed=" + editor + "&t=" + tit + "&e=" + exi;
                    window.location.href = vinculo;
                };
            }

            function DM02_NuevaLinea(paquete, objeto, editor) {
                var la = document.getElementById("DM02_Lado").value;
                var tit2 = document.getElementById("DM02_Titulo").value;
                var tit = tit2.replace(/&/g, "yy");
                var exi = document.getElementById("DM02_Exito").value;

                if (tit != "") {
                    var vinculo = "objeto_grabar_detalle.asp?op=2&p=" + paquete + "&o=" + objeto + "&ed=" + editor + "&t=" + tit + "&e=" + exi + "&la=" + la;
                    window.location.href = vinculo;
                };
            }

            function DM03_NuevaLinea(paquete, objeto, editor) {
                var tit2 = document.getElementById("DM03_Titulo").value;
                var tit = tit2.replace(/&/g, "yy");          

                if (tit != "") {
                    var vinculo = "objeto_grabar_detalle.asp?op=3&p=" + paquete + "&o=" + objeto + "&ed=" + editor + "&t=" + tit;
                    window.location.href = vinculo;
                };          
            }

            function DM04_NuevaLinea(paquete, objeto, editor) {
                var tit2 = document.getElementById("DM04_Titulo").value;
                var tit = tit2.replace(/&/g, "yy");          
                var num = document.getElementById("DM04_NumSerieLlave").value;

                if (tit != "") {
                    var vinculo = "objeto_grabar_detalle.asp?op=4&p=" + paquete + "&o=" + objeto + "&ed=" + editor + "&t=" + tit + "&num=" + num;
                    window.location.href = vinculo;
                };
            }

            function DM05_NuevaLinea(paquete, objeto, editor) {
                var tit2 = document.getElementById("DM05_Titulo").value;
                var tit = tit2.replace(/&/g, "yy");          

                if (tit != "") {
                    var vinculo = "objeto_grabar_detalle.asp?op=5&p=" + paquete + "&o=" + objeto + "&ed=" + editor + "&t=" + tit;
                    window.location.href = vinculo;
                };
            }

            function DM06_NuevaLinea(paquete, objeto, editor) {
                var prot = document.getElementById("DM06_Protagonista").value;

                if (prot != "") {
                    var vinculo = "objeto_grabar_detalle.asp?op=6&p=" + paquete + "&o=" + objeto + "&ed=" + editor + "&prot=" + prot;
                    window.location.href = vinculo;
                };
            }      

            function DM07_NuevaLinea(paquete, objeto, editor) {
                var idi = document.getElementById("DM07_Idiomaa").value;
                var audio = document.getElementById("DM07_Audio").value;
                var subt = document.getElementById("DM07_SubTitulos").value;

                var vinculo = "objeto_grabar_detalle.asp?op=7&p=" + paquete + "&o=" + objeto + "&ed=" + editor + "&idi=" + idi + "&aud=" + audio + "&sub=" + subt;
                window.location.href = vinculo;
            }

            function Filtro(cadena, titulo) {
                vinculo = "/forma/desktop/apps/discos/lista_paquetes.asp?f=" + cadena + "&t=" + titulo;            
                window.location.href = vinculo;
            }

            function obj_borrar_det(secuencia, editor, paquete, objeto) {
                var confirmacion = confirm("Desea borrar el detalle seleccionado?");
                var vinculo = "objeto_borrar_detalle.asp?s=" + secuencia + "&ed=" + editor + "&p=" + paquete + "&o=" + objeto;

                if (confirmacion) {     
                window.location.href = vinculo;
                };
            }

            function obj_borrar_prot(secuencia, editor, paquete, objeto) {
                var confirmacion = confirm("Desea borrar al protagonista seleccionado?");
                var vinculo = "objeto_borrar_protagonista.asp?s=" + secuencia + "&ed=" + editor + "&p=" + paquete + "&o=" + objeto;

                if (confirmacion) {     
                window.location.href = vinculo;
                };        
            }

            function obj_borrar_idioma(secuencia, editor, paquete, objeto) {
                var confirmacion = confirm("Desea borrar el idioma seleccionado?");
                var vinculo = "objeto_borrar_idioma.asp?s=" + secuencia + "&ed=" + editor + "&p=" + paquete + "&o=" + objeto;

                if (confirmacion) {     
                window.location.href = vinculo;
                };        
            }      

            mask(document.getElementById('AEdicion'), ['9999']);

            <% if tt("Editor") = "PE" or tt("Editor") = "VM" or tt("Editor") = "LI" then %>
                mask(document.getElementById('DuraPag'),  ['9999']);           
            <% end if %>
        </script>

        <% con.close: set con = nothing %> 
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>