<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    '
    ' Init()
    '

    dim cc, sqlString, Nuevo, Sistema, Nombre, Descripcion, ClaseApp, IndiceOrdenamiento, Icono, sBitacora, Externo

      '
    ' Funciones y Procedimientos
    '

    function limpiar(cadena)
        limpiar = Replace(cadena,"'","´")    
    end function

    Function ValorNull(Cadena)
        if isnull(Cadena) or (Len(Trim(Cadena)) = 0) then
            ValorNull = "NULL"
        else
            ValorNull = "'" & Cadena & "'"
        end if
    end function

    sub Actualizar()
        Sistema = Request.Form("Sistema")
        Proceso = Request.Form("Codigo")
        Nombre = Request.Form("Nombre")
        Action = Request.Form("Action")
        proActionParam = Request.Form("proActionParam")
        Activo = Request.Form("Activo")
        MenuItem = Request.Form("MenuItem")
        MenuIndice = Request.Form("MenuIndice")
        Root = Request.Form("Root")
        Snippet = Request.Form("Snippet")
        Shortcut = Request.Form("Shortcut")
        Movil = Request.Form("Movil")
        Icon = Request.Form("Icon")
        HPage = Request.Form("pHomePage")
        Externo = Request.Form("Externo")

        Nuevo = request.form("Nuevo")
        OrdenadoPor = Request.Form("ordenadoPor")

        if MenuItem = 0 then
            '
            ' Es un Menú CABECERA...
            ' El "Root" es EL SISTEMA PROPIETARIO...
            '
            Root = Sistema
        end if

        if Nuevo = "1" then
            sqlString = "INSERT INTO seg_Procesos(proSistema, proCodigo, proNombre, proActivo, proMenuItem, proMenuIndice, proIcon, proRoot, proAction, proHomePage, Snippet, Shortcut, Movil, Externo, proActionParam) " & _
                             "VALUES ('" & Sistema & "', '" & Proceso & "', '" & Nombre & "', " & Activo & ", " & MenuItem & ", " & MenuIndice & ", " & ValorNull(Icon) & ", '" & Root & "', " & ValorNull(Action) & ", " & HPage & ", " & ValorNull(Snippet) & ", " & Shortcut & ", " & Movil & ", " & Externo & ", " & ValorNull(proActionParam) & ");"
        else
            sqlString = "UPDATE seg_Procesos " & _
                           "SET proNombre = '"     & Nombre                     & "'," & _
                              " proActivo = "      & Activo                     & "," & _
                              " proMenuItem = "    & MenuItem                   & "," & _
                              " proMenuIndice = "  & MenuIndice                 & "," & _
                              " proIcon = "        & ValorNull(Icon)            & "," & _
                              " proRoot = '"       & Root                       & "'," & _
                              " proAction = "      & ValorNull(Action)          & "," & _
                              " proActionParam = " & ValorNull(proActionParam)  & "," & _
                              " proHomePage = "    & HPage                      & "," & _
                              " Snippet = "        & ValorNull(Snippet)         & "," & _
                              " Movil = "          & Movil                      & "," & _
                              " Externo = "        & Externo                    & "," & _
                              " Shortcut = "       & Shortcut                   & _
                        " WHERE (proSistema = '" & Sistema & "') " & _
                           "AND (proCodigo = '" & Proceso & "');"
        end if

        set cc = Server.CreateObject("ADODB.Connection")
        cc.open Application("Conn")
            cc.execute(sqlString)    
        cc.close: set cc = nothing

response.write sqlString        
    end sub

    '
    ' Main()
    '
    
    if Request.Cookies("usuario") = "" then
        Response.Redirect "../login.asp"      
    end if

    Actualizar

   Response.redirect "procesos.asp?s=" & Sistema & "&o=" & OrdenadoPor
%>