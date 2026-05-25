<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    '
    ' Parsear valores desde un WebAPI
    '

    Dim objSrvHTTP, res, con, t, sqlString, desde, final, tam, valor

    Set objSrvHTTP = Server.CreateObject ("Msxml2.ServerXMLHTTP.6.0")

    objSrvHTTP.open "GET", "https://openexchangerates.org/api/latest.json?app_id=555b27eb183f4776bdc273033635b722", false    
    objSrvHTTP.send

    Response.ContentType = "text/html"
    res = objSrvHTTP.responseText

    response.write RespuestaWebAPI

    '
    ' Ya tenemos los valores del WebAPI en "res"
    '

    set con = Server.CreateObject("ADODB.Connection")
    con.open Application("Conn")
    set t = con.Execute("SELECT Local, Simbolo, Formula " & _
                          "FROM seg_Cripto_NumParse_Locales " & _
                         "WHERE (Local <> 'NUM') " & _
                           "AND (Simbolo <> 'USD');")

        If Not (t.BOF Or t.EOF) Then
            Do
                desde = InStr(1, res, t("Simbolo")) + Len(t("Simbolo")) + 2
                final = InStr(desde, res, ",")
                valor = mid(res, desde, (final - desde))

                sqlString = "UPDATE seg_Cripto_NumParse_Locales" & _
                              " SET Formula = " & valor & _
                            " WHERE (Local = '" & t("Local") & "');"
                con.execute (sqlString)
               
                t.MoveNext
            Loop Until t.EOF
        End If

    t.close: set t = nothing
    con.close: set con = Nothing

    response.redirect "lista.asp"
%>




                                      

