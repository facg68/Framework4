<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->          
        <%
            dim cc, tt, Sistema

            Parametro = Request.QueryString("p")
            Sistema = Request.QueryString("s")

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")

            Function UsuariosAsignados(Parametro)
                set tt = cc.execute("SELECT COUNT(*) AS Cuantos FROM seg_Usuarios_Parametros WHERE (Parametro = '" & Parametro & "');")
                    UsuariosAsignados = tt("Cuantos")
                tt.close: set tt = nothing
            end Function
        %>
    </head>

    <body style="text-align: center;">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->     

        <img src='imagenes/none.png' style='border: none; width:5px; height: 30px;'>

        <p>
            <span style="font-size: 24pt; color: rgb(224, 62, 45);">
                <strong>No Se Puede Borrar la Variable</strong>
            </span>
        </p>        
        
        <hr>
        
        <p>
            <span style="font-size: 18pt; color: rgb(35, 111, 161);">
                La variable que está tratando de borrar tiene<br>
                asignada a uno o más usuarios.<br><br>
            
                Para eliminar esta variable, deben quitarse las asignaciones<br>
                que tenga y luego intentarlo nuevamente.
            </span>
        </p>

        <%
            if UsuariosAsignados(Parametro) = 0 then
                cc.execute("DELETE FROM seg_Parametros WHERE Parametro = '" & Parametro & "';") 
                response.redirect "variables.asp?s=" &  Sistema & "&o=" & request.QueryString("o")
            end if

            cc.close: set cc = nothing
        %>

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->       
    </body>
</html>