<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->        
        <%
            dim cc, tt, Rol, unico, ordenadoPor, Sistema, vinculo

            Rol = Request.QueryString("r")
            ordenadoPor = Request.QueryString("o")
            Sistema = Request.QueryString("s")
            unico = Request.QueryString("u")

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")

            Function TieneUsuarios(Rol)
                set tt = cc.execute("SELECT COUNT(*) AS Cuantos FROM seg_RolesUsuarios WHERE (CodigoRol = '" & Rol & "')")
                    TieneUsuarios = tt("Cuantos")
                tt.close: set tt = nothing
            end Function
        %>
    </head>

    <body plantilla="normal" reserva="165" style="text-align: center;">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->     

        <img src='imagenes/none.png' style='border: none; width:5px; height: 30px;'>

        <p>
            <span style="font-size: 24pt; color: rgb(224, 62, 45);">
                <strong>No Se Puede Borrar el Rol</strong>
            </span>
        </p>        
        
        <hr>
        
        <p>
            <span style="font-size: 18pt; color: rgb(35, 111, 161);">
                El rol que está tratando de borrar tiene<br>
                asignados a uno o más usuarios.<br><br>
            
                Para eliminar este rol, deben quitarse las asignaciones<br>
                que tenga y luego intentarlo nuevamente.
            </span>
        </p>

        <%
            if TieneUsuarios(Rol) = 0 then
                cc.execute("DELETE FROM seg_Roles WHERE (rolCodigo = '" & Rol & "')") 

                if unico = "1" then
                    vinculo = "roles_sys.asp?s=" & Sistema & "&o=" & OrdenadoPor
                else
                    vinculo = "roles.asp?s=" & Sistema & "r=" & Rol & "&o=" & OrdenadoPor
                end if

                response.redirect vinculo
            end if

            cc.close: set cc = nothing
        %>

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->   
    </body>
</html>