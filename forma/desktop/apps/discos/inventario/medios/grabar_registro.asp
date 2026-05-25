<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 

<%
    '
    ' Variables Globales
    '

    dim cc, tt, sqlString, secuencia
    dim ACompra, AEdicion, Titulo, Casa, Tienda, Precio, Descripcion
    dim VerComo, Carpeta, Paquete, Editor

    Paquete = Request.Form("Paquete")
    ACompra = Request.Form("ACompra")
    AEdicion = Request.Form("AEdicion")
    Titulo = LimpiarApostrofes(Request.Form("Titulo"))
    Precio = Request.Form("Precio")
    Tienda = Request.Form("Tienda")
    Casa = Request.Form("Casa")
    Descripcion = LimpiarApostrofes(Request.Form("Descripcion"))
    VerComo = Request.Form("VerComo")
    Carpeta = Request.Form("Carpeta")

    Amo = Year(Now())
    Editor = Request.Form("cboNuevoTipoObjeto")

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")

    '
    ' Funciones y procedimiemtos
    '

    Function SecuenciaObjeto(Tipo)
        sqlString = "select ISNULL(MAX(CAST(right(Objeto, 5) AS NUMERIC(5,0))),0) AS Valor " & _
                    "from discos_Objetos " & _
                    "where (usuario ='" & Request.Cookies("Usuario") & "') " & _
                    "and (Objeto Like '" & Tipo & "%');"

        set tt = cc.execute(sqlString)

        if not (tt.bof or tt.eof) then
            secuencia = cDbl(tt("Valor")) + 1
            SecuenciaObjeto = Tipo & RIGHT("00000" & secuencia, 5)
        end if

        tt.close: set tt = nothing
    end function

    function LimpiarApostrofes(valor)
        LimpiarApostrofes = Replace(valor,"'","´")
    end function

    sub ActualizarEncabezado()
        sqlString = "UPDATE discos_Paquetes " &  _
                    "SET ACompra = " & ACompra & ", " & _
                    " AEdicion = " & AEdicion & ", " & _
                    " Titulo = '" & Titulo & "', " & _
                    " Precio = " & Precio & ", " & _
                    " Tienda = '" & Tienda & "', " & _
                    " Casa = '" & Casa & "', " & _
                    " Descripcion = '" & Descripcion & "', " & _
                    " Carpeta = '" & Carpeta & "', " & _
                    " VerComo = " & VerComo & " " & _
                "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                    "AND (Paquete = '" & Paquete & "');"
        cc.execute(sqlString)
    end sub

    '
    ' Main()
    '

    ActualizarEncabezado

    Objeto = SecuenciaObjeto(Editor)            

    if Editor <> "*" then
        sqlString = "INSERT INTO discos_Objetos(Usuario, Paquete, Objeto, aEdicion, Titulo, Forma, Tipo, Clasificacion, FormatoPantalla, PlatOS, Editor) " & _
                    "VALUES('" & Request.Cookies("Usuario") & "', '" & Paquete & "', '" & Objeto & "',  " & Amo & ", 'Nuevo Objeto', '00000000', '00000000', '-', '00000000', '00000000', '" & Editor & "');"

        cc.execute(sqlString)

        response.redirect "editar_objeto.asp?p=" & Paquete & "&o=" & Objeto & "&e=" & Editor        
    else
        response.redirect "lista.asp"
    end if

    cc.close: set cc = nothing
%>