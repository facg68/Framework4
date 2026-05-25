<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim c, t, sqlString, Cuenta
    Cuenta = Request.QueryString("c")

    set c = Server.CreateObject("ADODB.Connection")
    c.open Application("Conn")

        c.execute("DELETE FROM pre_Cuentas " & _
                    "WHERE Usuario = '" & Request.Cookies("Usuario") & "' " & _
                    "AND Codigo = '" & Cuenta & "';")

    c.close: set c = nothing

    response.redirect "lista.asp"
%>    
