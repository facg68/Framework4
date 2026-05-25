<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            dim con, t, sqlString, Sistema, Version, Caracteristica, Descripcion

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")    
        %>
    </head>

    <body>
        <%
            Sistema = Request.Querystring("s")
            Version = Request.Querystring("v")
            Caracteristica = Request.Querystring("c")

            sqlString = "DELETE FROM seg_VersionesDetalles " & _
                        "WHERE (Sistema = '" & Sistema & "') "  & _
                        "AND (Version = '" & Version & "') "  & _
                        "AND (Caracteristica = " & Caracteristica & ");"

            con.execute sqlString
            con.close: set con = nothing

            response.redirect "editar_version.asp?s=" & Sistema & "&v=" & Version
        %>    
    </body>
</html>