<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            dim con, t, sqlString, Sistema, Version, Caracteristica, Descripcion, SolicitadoPor, FechaSolicitado

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")    

            function secuencia(Sistema, Version)
               sqlString = "SELECT ISNULL(MAX(Caracteristica), 0) + 1 AS Valor " & _
                             "FROM seg_VersionesDetalles " & _
                            "WHERE (Sistema = '" & Sistema & "') " & _
                              "AND (Version = '" & Version & "')"

                set t = con.execute(sqlString)
                    secuencia = t("Valor")                        
                t.close: set t = nothing
            end function

            function ValorNull(Valor)
                if isnull(valor) or (valor = "") then
                    ValorNull = "NULL"
                else
                    ValorNull = "'" & Valor & "'"
                end if
            end function

            function FechaServer(FechaFormulario)
                dim d, m, a, Fecha

                if len(FechaFormulario) <> 10 then
                    FechaServer = NULL
                else
                    Fecha = FechaFormulario
                    if len(trim(Fecha)) <> 10 then Fecha = NULL

                    d = right("00" & left(Fecha, 2), 2)
                    m = right("00" & mid(Fecha, 4, 2), 2)
                    a = right(Fecha, 4)

                    FechaServer = a & "-" & m & "-" & d
                end if
            end function    
        %>
    </head>

    <body>
        <%
            Sistema = Request.Querystring("s")
            Version = Request.Querystring("v")
            Caracteristica = secuencia(sistema, version)
            Descripcion = Request.Querystring("d")
            SolicitadoPor = Request.Querystring("sp")
            FechaSolicitado = FechaServer(Request.Querystring("fs"))

            sqlString = "INSERT INTO seg_VersionesDetalles(Sistema, Version, Caracteristica, Descripcion, SolicitadoPor, FechaSolicitado) " & _
                        "VALUES ('" & Sistema & "', '" & Version & "', " & Caracteristica & ",'" & Descripcion & "', " & ValorNull(SolicitadoPor) & ", " & ValorNull(FechaSolicitado) & ");"

            con.execute sqlString
            con.close: set con = nothing

            response.redirect "editar_version.asp?s=" & Sistema & "&v=" & Version        
        %>    
    </body>
</html>