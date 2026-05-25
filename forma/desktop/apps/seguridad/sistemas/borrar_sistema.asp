<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->          
        <%
            dim cc, tt, Sistema

            Sistema = Request.QueryString("s")

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")

            Function TieneRoles(Sistema)
                set tt = cc.execute("SELECT COUNT(*) AS Cuantos FROM seg_PermisosUsuarios WHERE (Sistema = '" & Sistema & "');")
                    TieneRoles = tt("Cuantos")
                tt.close: set tt = nothing
            end Function
        %>
    </head>

    <body plantilla="normal" reserva="165" style="text-align: center;">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->     

        <img src='imagenes/none.png' style='border: none; width:5px; height: 30px;'>

        <p>
            <span style="font-size: 24pt; color: rgb(224, 62, 45);">
                <strong>No Se Puede Borrar el Sistema</strong>
            </span>
        </p>        
        
        <hr>
        
        <p>
            <span style="font-size: 18pt; color: rgb(35, 111, 161);">
                El sistema que está tratando de borrar tiene procesos que se han<br>
                asignado a uno o más usuarios.<br><br>
            
                Para eliminar este sistema, deben quitarse las asignaciones<br>
                que tenga y luego intentarlo nuevamente.
            </span>
        </p>

        <%
            if TieneRoles(Sistema) = 0 then
                cc.execute("DELETE FROM seg_Sistemas WHERE sysCodigo = '" & Sistema & "';") 
                response.redirect "lista.asp?s=" & Sistema & "&op=" & request.QueryString("op")
            end if

            cc.close: set cc = nothing
        %>

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->      
    </body>
</html>