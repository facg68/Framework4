<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim con, t, sqlString, Sistema, Version, OrdenadoPor

    Sistema = Request.QueryString("s")
    Version = Request.QueryString("v")
    OrdenadoPor = Request.QueryString("o")

    set con = Server.CreateObject("ADODB.Connection")
    con.open Application("Conn")

    '
    ' Desactivamos cualquier versión con Estatus = 1... 
    ' Ahora será 2 (Histórica)
    '
    sqlString = "UPDATE seg_Versiones " & _
                   "SET Activa = 2 " & _
                 "WHERE (Sistema = '" & Sistema & "') " & _
                   "AND (Activa = 1);"

    con.execute(sqlString)

    '
    ' Ahora "promovemos" a ACTUAL a la versión solicitada...
    '

    sqlString = "UPDATE seg_Versiones " & _
                   "SET Activa = 1, " & _
                      " FechaActivacion = GETDATE() " & _
                 "WHERE (Sistema = '" & Sistema & "') " & _
                   "AND (Version = '" & Version & "');"
    
    con.execute(sqlString)

    response.redirect "versiones.asp?s=" & Sistema & "&v=" & Version & "&0=" & OrdenadoPor
%>
