<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <title>Cambiar Clave de Acceso del Usuario</title>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->    

        <%
            '
            ' Init()
            '
            dim cc, tt, sqlString

            sqlString = "SELECT * " & _
                        "FROM seg_Usuarios " & _
                        "WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';"

            if Request.Cookies("usuario") = "" then
                Response.Redirect "../login.asp"      
            end if

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")
            set tt = cc.execute(sqlString)

            '
            ' Funciones y Procedimientos
            '

            function hPage()
                dim t, cmdString, con

                cmdString = "SELECT usuHomePage FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';"

                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")
                    set t = con.execute(cmdString)
                        if not (t.bof or t.eof) then
                            if len(trim(t("usuHomePage"))) > 0 then
                                hPage = "/forma/desktop/apps/" & t("usuHomePage") & ".asp"
                            else
                                hPage = "/core/desktop.asp"
                            end if
                        else
                            hPage = Application("DefPage")
                        end if
                    t.close: set t = nothing        
                con.close: set con = nothing
            end function        
        %>  

        <style>
            body { overflow: hidden; }
        </style>                
    </head>

    <body>
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <br />

        <form id="formulario"  name="formulario" method="post" action="profile_pass_actualizar.asp">
            <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
                <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                    Cambiar Clave de Acceso
                </div>
                
                <div style="flex: 0 0 50%; text-align: right;">                    
                    <button class='form-btn azul normal'  type='button' onclick="volver()">Cancelar</button>
                    <button class='form-btn verde normal' type='button' onclick="grabar()">Actualizar</button>                    
                </div>
            </div>        

            <div class="main main-scroll">
                <div class="no-ver">
                  <input id="codigo" name="codigo" type="text" value="<%= Request.Cookies("Usuario") %>" />
                </div>           

                <div class="line">
                    <label class="label normal">Clave Actual</label>
                    <input class="field normal" id="password_actual" name="password_actual" type="password" required >
                </div>

                <div class="line">
                    <label class="label normal">Nueva Clave</label>
                    <input class="field normal" id="password_nuevo1" name="password_nuevo1" type="password" required >
                </div>

                <div class="line">
                    <label class="label normal">Verificar Nueva Clave</label>
                    <input class="field normal" id="password_nuevo2" name="password_nuevo2" type="password" required >
                </div>

                <div class="line">
                    <label class="label full">No puede usar el apóstrofe (') en una clave de acceso</label>
                </div>    
            </div>
        </form>
  
        <br /><br />   

        <script type="text/javascript">
            function volver() {
                var vinculo = "<%= hPage() %>";
                window.location.href = vinculo;
            }

            function grabar() {
                document.getElementById("formulario").submit();
            }

            mask(document.getElementById('usuFechaNacimiento'), ['99/99/9999']);
            mask(document.getElementById('snippetsOpacidad'),   ['999']);        
        </script>

        <%  
            tt.close: set tt = nothing
            cc.close: set cc = nothing    
        %>            

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>