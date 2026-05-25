<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            dim con, t, sqlString, Sistema, Version, Resumen, Obligatoria, OrdenadoPor
            dim cuantos, k, Campo1, Campo2,Campo3, Campo4, ValorCampo1, ValorCampo2, ValorCampo3, ValorCampo4

            Sistema = Request.Form("Sistema")
            Version = Request.Form("Version")
            Resumen = Request.Form("Resumen")
            Obligatoria = Request.Form("Obligatoria")
            OrdenadoPor = Request.Form("OrdenadoPor")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            function CuantoDetalle(Sistema, Version)
                dim sqlString

                sqlString = "SELECT COUNT(*) AS Cuantos " & _ 
                              "FROM seg_VersionesDetalles " & _
                             "WHERE (Sistema = '" & Sistema & "') " & _
                               "AND (Version = '" & Version & "');"

                set t = con.execute(sqlString)   
                    CuantoDetalle = t("Cuantos") 
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
            Cuantos = CuantoDetalle(Sistema, Version)        

            sqlString = "UPDATE seg_Versiones " & _
                           "SET Resumen = '" & Resumen & "', " & _
                              " Obligatoria = " & Obligatoria & _
                        " WHERE (Sistema = '" & Sistema & "') "  & _
                           "AND (Version = '" & Version & "');"
            
            con.execute(sqlString)

response.write sqlString & "<br /><br />"

            if Cuantos > 0 then
                for k = 1 to Cuantos
                    campo1 = "FORM_Caracteristica_" & k
                    campo2 = "FORM_Descripcion_" & k
                    campo3 = "FORM_SolicitadoPor_" & k
                    campo4 = "FORM_FechaSolicitado_" & k

                    ValorCampo1 = Request.Form(Campo1)
                    ValorCampo2 = Request.Form(Campo2)
                    ValorCampo3 = Request.Form(Campo3)
                    ValorCampo4 = FechaServer(Request.Form(Campo4))                    

                    sqlString = "UPDATE seg_VersionesDetalles " & _
                                   "SET Descripcion = '" & ValorCampo2 & "', " & _
                                      " SolicitadoPor = " & ValorNull(ValorCampo3) & ", " & _
                                      " FechaSolicitado = " & ValorNull(ValorCampo4) & _
                                " WHERE (Sistema = '" & Sistema & "') " & _
                                   "AND (Version = '" & Version & "') " & _
                                   "AND (Caracteristica = " & ValorCampo1 & ");"

                    con.execute(sqlString)
response.write sqlString & "<br /><br />"                    
                next
            end if

            response.redirect "versiones.asp?s=" & Sistema & "&v=" & Version & "&o=" & OrdenadoPor
        %>
    </body>
</html>