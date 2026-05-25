<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    '
    ' Init()
    '

    dim cc, tt, ptt, sqlString, usuario
    dim nombre, cargo, correo, fecha, telefono, estado, homepage, versaldos, IniciarSinEncabezado
    dim cargarSnippets, snippetsOpacidad, usuRandomWallpaper

    usuario = Request.Form("codigo")   

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")    
   
    '
    ' Funciones y Procedimientos
    '
    function TieneSnippets()
        sqlCommand = "SELECT CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END AS Cuantos " & _
                       "FROM dbo.seg_Usuarios_Snippets AS s " & _
                      "WHERE (codUsuario = '" & Usuario & "');"   

        set ptt = cc.execute(sqlCommand)  
            TieneSnippets = ptt("Cuantos") 
        ptt.close: set ptt = nothing                                   
    end function

    function TieneVariables()
        sqlCommand = "SELECT COUNT(*) AS Cuantas " & _
                       "FROM dbo.seg_Parametros AS p INNER JOIN dbo.seg_Usuarios_Parametros AS up " & _ 
                         "ON p.Parametro = up.Parametro AND p.Sistema = up.Sistema " & _
                      "WHERE (up.Usuario = '" & Usuario & "') AND (p.Exponer = 1);"    

        set ptt = cc.execute(sqlCommand)  
            TieneVariables = ptt("Cuantas") 
        ptt.close: set ptt = nothing                                   
    end function

    function TieneShortcuts()
        sqlCommand = "SELECT CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END AS Cuantos " & _
                       "FROM dbo.seg_PermisosUsuarios AS pu " & _
                 "INNER JOIN dbo.seg_Procesos AS pr " & _
                         "ON pu.Sistema = pr.proSistema " & _
                        "AND pu.Proceso = pr.proCodigo " & _
                      "WHERE (pu.Usuario = '" & Usuario & "') " & _
                        "AND (pr.Shortcut = 1);"

        set ptt = cc.execute(sqlCommand)  
            TieneShortcuts = ptt("Cuantos")
        ptt.close: set ptt = nothing     
    end function   

    function FechaServidor(FechaFormulario)
        dim d, m, a

        d = left(FechaFormulario, 2)
        m = mid(FechaFormulario, 4, 2)
        a = right(FechaFormulario, 4)

        FechaServidor = a & "-" & m & "-" & d
    end function

    function limpiar(cadena)
        limpiar = Replace(cadena,"'","´")    
    end function

    sub ActualizarUsuario()
        nombre = limpiar(request.form("usuNombre"))
        cargo = limpiar(request.form("usuCargo"))
        correo = limpiar(request.form("usuCorreo"))
        estado = request.form("usuEstado")

        fecha = request.form("usuFechaNacimiento")
        if len(trim(fecha)) = 10 then
            fecha = FechaServidor(fecha)
        else
            fecha = "2001-01-01"
        end if

        telefono = limpiar(request.form("usuTelefono"))
        homepage = request.form("usuHomePage")

        versaldos = request.form("usuversaldos")
        IniciarSinEncabezado = request.form("usuIniciarSinEncabezado")
        CargarSnippets = request.form("usuCargarSnippets")
        usuRandomWallpaper = request.form("usuRandomWallpaper")

        sqlString = "UPDATE seg_Usuarios " & _
                       "SET usuNombre = '" & nombre & "', " & _
                          " usuCargo = '" & cargo & "', " & _
                          " usuCorreo = '" & correo & "', " & _
                          " usuFechaNacimiento = '" & fecha & "', " & _
                          " usuEstado = " & Estado & ", " 

        if len(trim(homepage)) > 0 then
            sqlString = sqlString & "usuHomePage = '" & homepage & "', "
        else
            sqlString = sqlString & "usuHomePage = NULL, "
        end if

        if versaldos = 1 then
            sqlString = sqlString & "usuVerSaldos = 1, "
        else
            sqlString = sqlString & "usuVerSaldos = 0, "
        end if

        if IniciarSinEncabezado = 1 then
            sqlString = sqlString & "usuIniciarSinEncabezado = 1, "
        else
            sqlString = sqlString & "usuIniciarSinEncabezado = 0, "
        end if    

        if usuRandomWallpaper = 1 then
            sqlString = sqlString & "usuRandomWallpaper = 1, "
        else
            sqlString = sqlString & "usuRandomWallpaper = 0, "
        end if        

        if CargarSnippets = 1 then
            sqlString = sqlString & "usuCargarSnippets = 1, "
        else
            sqlString = sqlString & "usuCargarSnippets = 0, "
        end if        

        if TieneSnippets() = 1 then
            snippetsOpacidad = request.form("snippetsOpacidad")

            sqlString = sqlString & "snippetsOpacidad = " & snippetsOpacidad & ", "
        end if
                                                                             
        sqlString = sqlString & "usuTelefono = '" & telefono & "' " & _
                  " WHERE usuCodigo = '" & usuario & "';"

        cc.execute(sqlString)    
    end sub

    Sub ActualizarSnippets()
        dim t, tSQL, tValor

        if TieneSnippets() then
            set t =  cc.execute("SELECT codSistema, codProceso, snippet FROM seg_Usuarios_Snippets WHERE (codUsuario = '" & usuario & "');")

            do
                tValor = Request.Form(t("Snippet"))

                tSQL = "UPDATE seg_Usuarios_Snippets SET snippetActivo = "
                if tValor = 1 then
                    tSQL = tSQL & "1"
                else
                    tSQL = tSQL & "0"
                end if
                tSQL = tSQL & " WHERE (codUsuario = '" & usuario & "') " & _
                                "AND (Snippet = '" & t("Snippet") & "');"
                
                cc.execute(tSQL)

                t.MoveNext
            loop until t.eof

            t.close: set t = nothing
        end if
    end Sub

    Function QueSistema(Proceso)
        dim ss
        
        set ss = cc.execute("SELECT proSistema FROM dbo.seg_Procesos WHERE (proCodigo = '" & Proceso & "');")
            QueSistema = ss("proSistema")
        ss.close: set ss = nothing    
    End Function

    Sub InsertarShortCut(Proceso)
        dim sqlCommand

        sqlCommand = "INSERT INTO seg_Usuarios_Shortcuts(codUsuario, codSistema, codProceso) " & _
                           "VALUES ('" & usuario & "', '" & QueSistema(Proceso) & "', '" & Proceso & "');"

        cc.execute(sqlCommand)                              
    End Sub

    Sub ProcesarShortcuts()
        cc.execute("DELETE FROM seg_Usuarios_Shortcuts WHERE codUsuario = '" & usuario & "'")
        set t =  cc.execute("exec dbo.seg_Shortcuts_Activos '" & usuario & "';")

        do
            if Request.Form(t("Proceso")) = 1 then
                InsertarShortCut t("Proceso")
            end if

            t.MoveNext
        loop until t.eof

        t.close: set t = nothing
    end Sub

    '
    ' Main()
    '

    ActualizarUsuario 
    if TieneShortcuts() = 1 then ProcesarShortcuts 
    if TieneSnippets = 1 then ActualizarSnippets 

    cc.close: set cc = nothing

    Response.redirect "lista.asp?o=" & ordenadoPor
%>