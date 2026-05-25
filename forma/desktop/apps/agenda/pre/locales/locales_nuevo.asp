<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim c, t, sqlString
    dim Local, Nombre, Simbolo, Formula

    Local = Request.Form("local")
    Nombre = Request.Form("Nombre")
    Simbolo = Request.Form("Simbolo")
    Formula = Request.Form("Formula")

    if Formula = "" then Formula ="0.00"    

    sqlString = "INSERT INTO seg_Cripto_NumParse_Locales(Local, Nombre, Simbolo, Formula) " & _
                "VALUES('" & Local & "', '" & Nombre & "', '" & Simbolo & "', " & Formula & ");"

    set c = Server.CreateObject("ADODB.Connection")
    c.open Application("Conn")
        c.execute(sqlString)
    c.close: set c = nothing

    response.redirect "lista.asp"
%>