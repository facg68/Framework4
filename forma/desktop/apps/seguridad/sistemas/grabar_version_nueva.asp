<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    '
    ' Init()
    '

    dim cc, sqlString, Version, Sistema, Resumen, Obligatoria, FechaActivacion

      '
    ' Funciones y Procedimientos
    '

    function limpiar(cadena)
        limpiar = Replace(cadena,"'","´")    
    end function

    '
    ' Main()
    '
    
    if Request.Cookies("usuario") = "" then
        Response.Redirect "/apps/desktop/login.asp"      
    end if

    Version = Request.Form("Version")
    Sistema = Request.Form("Sistema")
    Resumen = limpiar(Request.Form("Resumen"))
    Obligatoria = Request.Form("Obligatoria")


    sqlString = "INSERT INTO seg_Versiones(Version, Sistema, Resumen, Obligatoria, Activa) " & _
                "VALUES ('" & Version & "', '" & Sistema & "', '" & Resumen & "', " & Obligatoria & ", 0);"

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")
        cc.execute(sqlString)    
    cc.close: set cc = nothing

   Response.redirect "editar_version.asp?s=" & Sistema & "&v=" & Version & "&o=" & Request.Form("ordenadoPor")
%>