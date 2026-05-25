<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    '
    ' Init
    '

    dim con, sqlstring, estatus, codigo, nombre, descripcion

    set con = Server.CreateObject("ADODB.Connection")
    con.open Application("Conn")

    '
    ' Functions()
    '

    function LimpiarApostrofes(valor)
        LimpiarApostrofes = Replace(valor,"'","´")
    end function

    function NullValue(Cadena)
        if isNull(Cadena) then
            NullValue = "NULL"
        else 
            NullValue = "'" & Cadena & "'"
        end if   
    end function

    function FechaForm2Server(FechaForm)
        dim d, m, a, h, mm

        if isnull(FechaForm) then
            NullValue = "NULL"
        else
            d = left(FechaForm, 2)
            m = mid(FechaForm, 4, 2)
            a = mid(FechaForm, 7, 4)
            h = mid(FechaForm, 12, 2)
            mm = mid(FechaForm, 15, 2)

            FechaForm2Server = "'" & a & "-" & m & "-" & d & " " & h & ":" & mm & "'"
        end if
    end function

    '
    ' Main()
    '
    estatus = Request.Form("estatus")
    codigo = Request.Form("codigo")
    nombre =  NullValue(Request.Form("nombre"))
    descripcion = NullValue(Request.Form("descripcion"))

    if estatus = 0 then
        sqlstring = "UPDATE seg_Anuncios_Pantallas " & _
                       "SET Nombre = " & nombre & ", " & _
                       "Descripcion = " & descripcion & " " & _
                   " WHERE pantalla = '" & codigo & "';"
    else
        sqlstring = "INSERT INTO seg_Anuncios_Pantallas(Pantalla, Nombre, Descripcion) " & _
                    "VALUES('" & Codigo & "', " & Nombre & ", " & Descripcion & ");"
    end If 

    con.execute (sqlString)
    con.close: set con = nothing

    response.redirect "pantallas.asp"

response.write sqlString    
%>