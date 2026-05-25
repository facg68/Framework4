<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function QueContacto(Secuencia)
                dim zCon, zTab

                set zCon = Server.CreateObject("ADODB.Connection")
                zCon.open Application("Conn")
                set zTab = zCon.execute("SELECT Codigo FROM con_Contactos_Adjuntos WHERE Secuencia = " & Secuencia & ";")

                    QueContacto = zTab("Codigo")

                ztab.close: set ztab= nothing
                zCon.close: set zCon = nothing
            end function

            sub borrarArchivo(Secuencia)
                dim zCon, zTab, filesPath, fOriginal

                set zCon = Server.CreateObject("ADODB.Connection")
                zCon.open Application("Conn")
                set zTab = zCon.execute("SELECT Nombre, Extension FROM con_Contactos_Adjuntos WHERE Secuencia = " & Secuencia & ";")

                    filesPath =lcase(Server.MapPath(lcase(Request.Cookies("usuPath")) & "/adjuntos") )
                    fOriginal = filesPath & "\" & zTab("Nombre") & "." & zTab("Extension")

                    Set fso = CreateObject("Scripting.FileSystemObject")
                        fso.DeleteFile fOriginal
                    set fso = nothing    

                ztab.close: set ztab= nothing
                zCon.close: set zCon = nothing
            end sub
        %>
    </head>

    <body>
        <% 
            dim c, sqlString, sec, cont

            usu = Request.Cookies("Usuario")
            objeto = Request.QueryString("s")
            cont = QueContacto(objeto)
            
            sqlString = "DELETE FROM con_Contactos_Adjuntos " & _
                        "WHERE (Secuencia = '" & objeto & "');"

            borrarArchivo objeto

            set c = Server.CreateObject("ADODB.Connection")
            c.open Application("Conn")
                c.execute sqlString
            c.close: set c = nothing

            response.redirect "cont_editar.asp?con=" & cont & "&tt=" & Request.QueryString("tt")
        %>
    </body>
</html>