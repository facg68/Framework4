<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 

<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->   

        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0120"
            SysLockOut
            

            dim con, t, p, sqlString, cbox
            dim Local, Nombre, Simbolo, MonedaEntera, MonedaEnteraUnica, MonedaFraccionada, MonedaFraccionadaUnica, Idioma, NombreListas, Formula

            Local = Request.QueryString("l")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            sqlString = "SELECT Nombre, Simbolo, MonedaEntera, MonedaEnteraUnica, MonedaFraccionada, MonedaFraccionadaUnica, Idioma, NombreListas, Formula " & _
                        "FROM seg_Cripto_NumParse_Locales " & _
                        "WHERE Local = '" & local & "';"

            set t = con.execute(sqlString)
                Nombre = t("Nombre") 
                Simbolo = t("Simbolo")
                MonedaEntera = t("MonedaEntera")
                MonedaEnteraUnica = t("MonedaEnteraUnica")
                MonedaFraccionada = t("MonedaFraccionada")
                MonedaFraccionadaUnica = t("MonedaFraccionadaUnica")
                Idioma = t("Idioma")
                NombreListas = t("NombreListas")
                Formula = t("Formula")
            t.close: set t = nothing            
        %>        

        <style>
            td {
                font-size: 14px;
                font-family: Ruda;
                padding: 3px;
            }
        </style>                
    </head>

    <body plantilla="normal" reserva="165">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <br />

        <form name="form_transaccion" id="form_transaccion" method="post" action="locales_grabar.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Editar Local: <%= Nombre %>                
                </div>
                
                <div style="flex: 0 0 50%; text-align: right;">
                    <button type="button" class="form-btn normal verde" onclick="submit()">
                        Actualizar
                    </button>

                    &nbsp;

                    <button type="button" class="form-btn normal rojo" onclick="AbrirVinculo('lista.asp')">
                        Cancelar
                    </button>  
                </div>
            </div>        

            <div class="main main-scroll"> 
                <div class="no-ver">
                    <input id="ordenamiento" name="ordenamiento" value="<%= oParan %>">
                    <input id="local" name="local" type="text" value="<%= local %>" />
                </div>

                <!-- Campos -->

                    <div class="line">
                        <label class="label xxl">Codigo</label>
                        <input class="field small" id="local" name="txtlocal" type="text" value="<%= local %>" disabled />
                    </div>

                    <div class="line">
                        <label class="label xxl">Simbolo</label>
                        <input class="field tiny" name="Simbolo" id="Simbolo" type="text" value="<%= Simbolo %>" >
                    </div>

                    <div class="line">
                        <label class="label xxl">Nombre</label>
                        <input class="field large" name="Nombre" id="Nombre" type="text" value="<%= Nombre %>" >
                    </div>

                    <div class="line">
                        <label class="label xxl">Moneda Entera</label>
                        <input class="field small" id="MonedaEntera" name="MonedaEntera" type="text" value="<%= MonedaEntera %>" />
                        <span style="font-size; 12px; color:rgb(71, 71, 71);">&nbsp;&nbsp;(ej. Dólares)</span>
                    </div>

                    <div class="line">
                        <label class="label xxl">Moneda Fraccionada</label>
                        <input class="field small" id="MonedaFraccionada" name="MonedaFraccionada" type="text" value="<%= MonedaFraccionada %>" />
                        <span style="font-size; 12px; color:rgb(71, 71, 71);">&nbsp;&nbsp;(ej. Centavos)</span>
                    </div>    

                    <div class="line">
                        <label class="label xxl">Moneda Fraccionada Única</label>
                        <input class="field small" id="MonedaFraccionadaUnica" name="MonedaFraccionadaUnica" type="text" value="<%= MonedaFraccionadaUnica %>" />
                        <span style="font-size; 12px; color:rgb(71, 71, 71);">&nbsp;&nbsp;(ej. Centavo)</span>
                    </div>                       

                    <div class="line">
                        <label class="label xxl">Idioma</label>

                        <select class="field normal" name="Idioma" id="Idioma" >
                            <%
                                set cbox = con.execute("SELECT Idioma, NombreIdioma FROM seg_Cripto_NumParse_Misc ORDER BY NombreIdioma;")

                                if not (cbox.bof or cbox.eof) then
                                    Do
                                        response.write "<option value='" & cbox("Idioma") & "' "
                                        if Idioma = cbox("Idioma") then 
                                            response.write " selected='selected'" 
                                        end if
                                        response.write ">" & cbox("NombreIdioma") & "</option>"

                                        cbox.MoveNext
                                    Loop Until cbox.eof
                                end if

                                cbox.close: set cbox = nothing
                            %>
                        </select>   
                    </div>

                   <div class="line">
                        <label class="label xxl">Nombre en Lista</label>
                        <input class="field normal" id="NombreListas" name="NombreListas" type="text" value="<%= NombreListas %>" />
                        <span style="font-size; 12px; color:rgb(71, 71, 71);">&nbsp;&nbsp;(ej. Dólares, Euros, etc.)</span>
                    </div>    

                    <div class="line">
                        <label class="label xxl">Formula</label>
                        <input class="field normal" name="Formula" id="Formula" type="number" value="<%= Formula %>" >
                        <span style="font-size; 12px; color:rgb(71, 71, 71);">&nbsp;&nbsp;(Equivalencia con 1.00 USD)</span>
                    </div>

                <!-- Fin de Campos -->
            </div>    
        </form>

        <br /><br />

        <script type="text/javascript">
            function submit(){
                document.getElementById("form_transaccion").submit(); 
            }

            function AbrirVinculo(vinculo) {
                window.location.href = vinculo;
            }

            mask(document.getElementById('Simbolo'), ['____________']);                
        </script>

        <% con.close: set con = nothing %>
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>