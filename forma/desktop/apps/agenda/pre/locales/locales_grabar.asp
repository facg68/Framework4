<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim cc, tt, sqlString, Local

    Local = Request.Form("local")

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")

        sqlString = "UPDATE seg_Cripto_NumParse_Locales " & _
                       "SET Nombre = '" & Request.Form("Nombre") & "'," & _
                          " Simbolo = '" & Request.Form("Simbolo") & "'," & _
                          " MonedaEntera = '" & Request.Form("MonedaEntera") & "'," & _
                          " MonedaEnteraUnica = '" & Request.Form("MonedaEnteraUnica") & "'," & _
                          " MonedaFraccionada = '" & Request.Form("MonedaFraccionada") & "'," & _
                          " MonedaFraccionadaUnica = '" & Request.Form("MonedaFraccionadaUnica") & "'," & _
                          " Idioma = '" & Request.Form("Idioma") & "'," & _
                          " NombreListas = '" & Request.Form("NombreListas") & "'," & _
                          " Formula = " & Request.Form("Formula") & _
                    " WHERE Local = '" & Local & "';"

        cc.execute(sqlString)

    cc.close: set cc = nothing

    response.redirect "lista.asp"
%>