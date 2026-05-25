<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    '
    ' Esta pagina graba la transaccion en la base de datos
    '

    function LimpiarApostrofes(valor)
        LimpiarApostrofes = Replace(valor,"'","´")
    end function

    '
    ' Main()
    '

    dim con, sqlString, nota, vinculo, NotaPre, NotaDonde
    dim usu, pre, llave, mDestino, mCambio

    '
    ' Leemos el formulario
    '
    Nota = LimpiarApostrofes(Request.form("txtNota"))
    NotaPre = Request.Form("NotaPre")
    NotaDonde = Request.Form("NotaDonde")
    llave = Request.form("Llave")
    usu = Request.Cookies("Usuario")
    pre = Request.Form("txtPre")

    dia = Request.Form("dia")
    ver = Request.Form("ver")
    tipo = Request.Form("tipo")
    estatus = Request.Form("estatus")
    ordenado = Request.Form("ordenado")    

    '
    ' Creamos la cadena de SQL
    '
    sqlString = "UPDATE pre_Presupuesto_Detalles " & _
                "SET Nota = '" & Nota & "'," & _
                   " NotaPre = " & NotaPre & ", " & _
                   " NotaDonde = '" & NotaDonde & "' " & _
                "WHERE (Llave = " & Llave & ");"

    '
    ' Ejecutamos el comando y volvemos al presupuesto
    '

    set con = server.CreateObject("ADODB.Connection")
    con.open Application("Conn")
        con.execute(sqlString)
    con.close: set con = nothing

    response.redirect "pre_det_editar.asp"
%>        