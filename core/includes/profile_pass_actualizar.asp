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

            dim cc, tt, sqlString, res
            dim nombre, cargo, correo, fecha, telefono

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")   

            if Request.Cookies("usuario") = "" then
                Response.Redirect "../login.asp"      
            end if

            '
            ' Funciones y Procedimientos
            '

            function limpiar(cadena)
                limpiar = Replace(cadena,"'","´")    
            end function

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
    </head>

    <body>
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                No se pudo Actualizar la Clave de Usuario!
            </div>
            
            <div style="flex: 0 0 50%; text-align: right;">
                <button class='form-btn verde normal' type='button' onclick="volver()">Volver</button>
            </div>
        </div>        

        <div class="main main-scroll">
            <br />

            Se encontraron errores en el proceso.<br /><br />

            <%
                actual   = limpiar(request.form("password_actual"))
                nuevo_01 = limpiar(request.form("password_nuevo1"))
                nuevo_02 = limpiar(request.form("password_nuevo2"))

                sqlString = "SELECT dbo.Cripto_VerifActPass('" & Request.Cookies("Usuario") & "', '" & actual & "', '" & nuevo_01 & "', '" & nuevo_02 & "') AS Res;"

                set tt = cc.execute(sqlString)
                    res = tt("res")
                tt.close: set tt = nothing

                if res <> 0 then
                    select case Res
                        case 1: response.write "La Clave Actual del Usuario no es correcta. No se puede realizar el proceso."
                        case 2: response.write "Las Claves Nuevas no Coinciden. No se puede realizar el proceso."
                    end select    
                else
                    '
                    ' El servidor reporta que se puede cambiar el password...
                    '
                    sqlString = "exec dbo.seg_pa_ActualizarClaveUsuario '" & Request.Cookies("Usuario") & "','" & nuevo_01 & "'"
                    cc.execute(sqlString)

                    Response.redirect hPage()
                end if
            %>                                                     

            <br /><br />

            Vuelva a la pantalla anterior e intente nuevamente
            
            <br />
        </div>
  
        <br /><br />   

        <script type="text/javascript">
            function volver() {
                var vinculo = "<%= hPage() %>";
                window.location.href = vinculo;
            }        
        </script>

        <% cc.close: set cc = nothing %> 
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>