<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <title>Editar Lista</title>    
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->             
        <%
            dim con, t, sqlString, cbox, c, tot_Precio, tot_Cambio

            dim Usuario, Codigo, Nombre, Descripcion, Cuenta, Monto, Contacto, PrecioOriginal
            dim PrecioFinal, MultiPrecio, Grupo, Categoria, VerListaEnInforme

            dim cuantos, nombreLista, MonedaOrigen, MonedaDestino
            dim Secuencia, Item, Precio, Fecha, t1


            thisSystem = "agenda"
            thisProcess = "agenda.0110"
            SysLockOut


            Usuario = Request.Cookies("Usuario")
            Codigo = Request.QueryString("l")
            cuantos = 0      

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")


            function LocalMonetarioUsuario()
                dim fcon, f

                set fcon = Server.CreateObject("ADODB.Connection")
                fcon.Open Application("Conn")
                set f = fcon.execute("SELECT isnull(usuLocal, 'US') AS usuLocal FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';")

                LocalMonetarioUsuario = f("usuLocal")
                
                f.close: set f = nothing
                fcon.close: set fcon = nothing
            end function

            function FechaServer(FechaForm)
                dim d, m, a

                if FechaForm = "" then 
                    FechaServer = NULL
                else
                    d = left(FechaForm, 2)
                    m = mid(FechaForm, 4, 2)
                    a = right(FechaForm, 4)

                    FechaServer = a & "-" & right("00" & m, 2) & "-" & right("00" & d, 2)
                end if
            end function   

            function FechaFormulario(FechaServer)
                dim d, m, a

                if FechaServer <> "" then      
                    d = right("00" & right(FechaServer, 2), 2)
                    m = right("00" & mid(FechaServer, 6, 2), 2)
                    a = left(FechaServer, 4)

                    FechaFormulario = d & "/" & m & "/" & a
                end if
            end function    
        %>   

        <style>
            .fila {
                display: flex;
                align-items: center;
                gap: 10px;
            }

            .col1 {
                white-space: nowrap;
                font-weight: bold;
                min-width: fit-content;
            }

            .col2 { flex: 0 0 10%; }
            .col3 { flex: 0 0 15%; }
            .col4 { flex: 1; }
            .col5 { flex: 0 0 60px; }

            a.linea, a.linea:link, a.linea:visited,
            a.linea:focus, a.linea:hover, 
            a.linea:active { color: black; }                
        </style>           
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            sqlString = "SELECT Usuario, Codigo, Nombre, Descripcion, Cuenta, Monto, Contacto, PrecioOriginal, PrecioFinal, MultiPrecio, Grupo, Categoria, VerListaEnInforme " & _
                        "FROM pre_Listas_Encabezado as e " & _
                        "WHERE (Usuario = '" & Usuario & "') " & _
                        "AND (Codigo = '" & Codigo & "');"

            set t = con.execute(sqlString)
                Nombre = t("Nombre")
                Descripcion =  t("Descripcion")
                Cuenta = t("Cuenta")
                Monto = t("Monto")
                Contacto = t("Contacto")
                PrecioOriginal = t("PrecioOriginal") 
                PrecioFinal = t("PrecioFinal")
                MultiPrecio = t("MultiPrecio")
                Grupo = t("Grupo")
                Categoria = t("Categoria")
                VerListaEnInforme = t("VerListaEnInforme")
            t.close: set t = nothing
        %>   

        <br />

        <form name="form_lista" id="form_lista" method="post" action="listas_actualizar.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Editar Lista <%= Nombre %>                
                </div>
                
                <div style="flex: 0 0 50%; text-align: right;">
                    <a onclick='submit()'>
                        <button type='button' class='form-btn azul' style='width: 100px; font-size: 16px; color: white;'>Actualizar</button>
                    </a>
                    
                    &nbsp;&nbsp;

                    <a href='lista.asp'>
                        <button type='button' class='form-btn rojo' style='width: 100px; font-size: 16px; color: white;'>Cancelar</button>
                    </a>                  
                </div>
            </div>        

            <div class="main main-scroll"> 
                <div class="no-ver">
                    <input id="Codigo"      name="Codigo"       value="<%= Codigo %>">
                    <input id="Usuario"     name="Usuario"      value="<%= Usuario %>">
                    <input id="MultiPrecio" name="MultiPrecio"  value="<%= MultiPrecio %>">     
                </div>

                <!-- Campos -->

                    <div class="line">
                        <label class="label normal">Nombre</label>
                        <input class="field xxl" name="nombre" id="nombre" type="text" value="<%= nombre %>">
                    </div>

                    <div class="line">
                        <label class="label normal">Descripción</label>
                        <input class="field xxl" name="descripcion" id="descripcion" type="text" value="<%= descripcion %>">
                    </div>

                    <div class="line">
                        <label class="label normal">Contacto</label>

                        <select class="field xxl" name="contacto" id="contacto" >
                            <%
                                sqlString = "SELECT Codigo, PrimerNombre + iif(PrimerApellido <> '', ' ' + PrimerApellido, '') AS NombreComtacto " & _
                                                "FROM con_Contactos " & _
                                                "WHERE usuario = '" & Usuario & "' " & _
                                                "AND visible = '1' " &  _
                                            "ORDER BY NombreComtacto;"

                                set cbox = con.execute(sqlString)

                                response.write "<option value=''"
                                if Contacto = "" then response.write " selected"
                                response.write ">&nbsp;</option>"

                                if not (cbox.bof or cbox.eof) then
                                    Do
                                        response.write "<option value='" & cbox("Codigo") & "' "
                                            if Contacto = cbox("Codigo") then 
                                                response.write " selected" 
                                            end if
                                        response.write ">" & cbox("NombreComtacto") & "</option>"

                                        cbox.MoveNext
                                    Loop Until cbox.eof
                                end if

                                cbox.close: set cbox = nothing
                            %>
                        </select>   
                    </div>

                    <div class="line">
                        <label class="label normal">Tipo</label>

                        <select class="field normal" name="Cuenta" id="Cuenta" >
                            <option value="0" <% if Cuenta = "0" then response.write " selected" %>>Es una lista</option>
                            <option value="1" <% if Cuenta = "1" then response.write " selected" %>>Es una Cuenta</option>
                        </select>                          
                    </div>

                    <div class="line">
                        <label class="label normal">Categoria</label>

                        <select class="field normal" name="categoria" id="categoria" >
                            <%
                                sqlString = "SELECT Codigo, Nombre " & _
                                            "FROM dbo.pre_Cuentas_Categorias AS c " & _
                                            "ORDER BY Nombre;"

                                set cbox = con.execute(sqlString)

                                if not (cbox.bof or cbox.eof) then
                                    Do
                                        response.write "<option value='" & cbox("Codigo") & "' "
                                            if categoria = cbox("Codigo") then 
                                                response.write " selected" 
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
                        <label class="label normal">Grupo</label>

                        <select class="field normal" name="grupo" id="grupo" >
                            <option value="A" <% if grupo = "A" then response.write " selected" %>>Activa</option>
                            <option value="W" <% if grupo = "W" then response.write " selected" %>>En Espera</option>
                            <option value="S" <% if grupo = "S" then response.write " selected" %>>Archivada</option>                
                        </select>                         
                    </div> 

                    <div class="line">
                        <label class="label normal">Local Original</label>

                        <select class="field normal" name="precioOriginal" id="precioOriginal" onChange="verificarPrecios()">
                            <%
                                sqlString = "SELECT [Local], Simbolo + '  (' + NombreListas + ')' AS NombreLocal " & _
                                                "FROM seg_Cripto_NumParse_Locales " & _
                                                "WHERE [Local] <> 'NUM' " & _
                                            "ORDER BY NombreLocal ASC;"

                                set cbox = con.execute(sqlString)

                                if not (cbox.bof or cbox.eof) then
                                    Do
                                        response.write "<option value='" & cbox("Local") & "' "
                                            if precioOriginal = cbox("Local") then 
                                                response.write " selected" 
                                            end if
                                        response.write ">" & cbox("NombreLocal") & "</option>"

                                        cbox.MoveNext
                                    Loop Until cbox.eof
                                end if

                                cbox.close: set cbox = nothing
                            %>
                        </select>   
                    </div>   

                    <div class="line">
                        <label class="label normal">Local Destino</label>

                        <select class="field normal" name="precioFinal" id="precioFinal" onChange="verificarPrecios()">
                            <%
                                sqlString = "SELECT [Local], Simbolo + '  (' + NombreListas + ')' AS NombreLocal " & _
                                                "FROM seg_Cripto_NumParse_Locales " & _
                                                "WHERE [Local] <> 'NUM' " & _
                                            "ORDER BY NombreLocal ASC;"

                                set cbox = con.execute(sqlString)

                                if not (cbox.bof or cbox.eof) then
                                    Do
                                        response.write "<option value='" & cbox("Local") & "' "
                                            if precioFinal = cbox("Local") then 
                                                response.write " selected" 
                                            end if
                                        response.write ">" & cbox("NombreLocal") & "</option>"

                                        cbox.MoveNext
                                    Loop Until cbox.eof
                                end if

                                cbox.close: set cbox = nothing
                            %>
                        </select>   
                    </div>                       

                    <div class="line">
                        <label class="label normal">Ver en Informe</label>

                        <select class="field normal" name="VerListaEnInforme" id="VerListaEnInforme" >
                            <option value="0" <% if VerListaEnInforme = "0" then response.write " selected" %>>Ocultar en Informe</option>
                            <option value="1" <% if VerListaEnInforme = "1" then response.write " selected" %>>Ver en Informe</option>
                        </select>  
                    </div> 

                <!-- Fin de los Campos -->                   
            </div>    
        </form>

        <br /><br />   

        <script type="text/javascript">
            function submit(){
                document.getElementById("form_lista").submit(); 
            }    

            function verificarPrecios(){
                var p1 = document.getElementById("precioOriginal").value;
                var p2 = document.getElementById("precioFinal").value;

                if (p1 == p2) {
                    document.getElementById("MultiPrecio").value = 0;
                }
                else {
                    document.getElementById("MultiPrecio").value = 1;
                }
            }            
        </script>

        <% con.close: set con = nothing %>    
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>