<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function FechaServer(FechaForm)
                dim d, m, a

                if FechaForm = "" then 
                    FechaServer = "NULL"
                else
                    d = left(FechaForm, 2)
                    m = mid(FechaForm, 4, 2)
                    a = right(FechaForm, 4)

                    FechaServer = a & "-" & right("00" & m, 2) & "-" & right("00" & d, 2)
                    FechaServer = "'" & FechaServer & "'"
                end if
            end function           
        %>
    </head>

    <body>
        <%
            dim c, t, sqlString, Cuenta, MultiPrecio

            set c = Server.CreateObject("ADODB.Connection")
            c.open Application("Conn")

                Cuenta = Request.Form("Cuenta")
                MultiPrecio = Request.form("MultiPrecio")

                '
                ' Parte 1: Actualizar el Encabezado de la Lista
                '

                sqlString = "UPDATE pre_Listas_Encabezado " & _
                            "SET Nombre = '" & Request.Form("nombre") & "', " & _
                            " Descripcion = '" & Request.Form("descripcion") & "', " & _
                            " Cuenta = " & Request.Form("Cuenta") & ", " & _
                            " Contacto = '" & Request.Form("contacto") & "', " & _
                            " PrecioOriginal = '" & Request.Form("precioOriginal") & "', " & _
                            " PrecioFinal = '" & Request.Form("precioFinal") & "', " & _
                            " MultiPrecio = " & Request.form("MultiPrecio") & ", " & _
                            " Grupo = '" & Request.Form("grupo") & "', " & _
                            " Categoria = '" & Request.Form("categoria") & "', " & _
                            " VerListaEnInforme = " & Request.Form("VerListaEnInforme") & _
                        " WHERE (Usuario = '" & Request.Form("Usuario") & "') " & _
                            "AND (Codigo = '" & Request.Form("Codigo") & "');"

                c.execute(sqlString) 
            c.close: set c = nothing

            response.redirect "lista.asp"
        %>    
    </body>
</html>


