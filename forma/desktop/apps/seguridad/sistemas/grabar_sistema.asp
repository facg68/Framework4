<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    '
    ' Init()
    '

    dim cc, sqlString, Nuevo, Sistema, Nombre, Descripcion, ClaseApp, IndiceOrdenamiento, Icono, sBitacora

    '
    ' Funciones y Procedimientos
    '

    function TieneVariables(Sistema)
        dim ptt, con, sqlCommand

        set con = Server.CreateObject("ADODB.Connection")
        con.open Application("Conn")

            sqlCommand = "SELECT s.sysCodigo, ISNULL(P.Cuantos, 0) AS Variables " & _
                        "FROM dbo.seg_Sistemas AS s " & _
                "LEFT OUTER JOIN (SELECT Sistema, COUNT(Parametro) AS Cuantos " & _
                                "FROM dbo.seg_Parametros " & _
                            "GROUP BY Sistema) AS P " & _
                            "ON s.sysCodigo = P.Sistema " & _
                        "WHERE (s.sysCodigo = '" & Sistema & "');"                                    

            set ptt = con.execute(sqlCommand)  
                if not (ptt.bof or ptt.eof) then
                    TieneVariables = ptt("Variables") 
                else
                    TieneVariables = 0
                end if                
            ptt.close: set ptt = nothing

        con.close: set con = nothing
    end function     

    sub ProcesarVariables(Sistema)
        dim ptt, con, sqlCommand, sqlTable, valor

        set con = Server.CreateObject("ADODB.Connection")
        con.open Application("Conn")        
        
            sqlTable = "SELECT Sistema, Parametro, TipoParametro, Sistema + '__' + Parametro AS Variable " & _
                         "FROM dbo.seg_Parametros AS p " & _
                        "WHERE (Sistema = '" & Sistema & "') " & _
                     "ORDER BY Descripcion;"

            set ptt =  con.execute(sqlTable)
                if not (ptt.bof or ptt.eof) then
                    do
                        valor = Request.Form(ptt("Variable")) 

                        sqlCommand = "UPDATE seg_Parametros " & _
                                        "SET ValorDefault = '" & valor & "' " & _
                                    "WHERE (Sistema = '" & ptt("Sistema") & "') " & _
                                        "AND (Parametro = '" & ptt("Parametro") & "');"

                        con.execute(sqlCommand)
                        
                        if ptt("TipoParametro") = 4 then 
                            ProcesarLista ptt("Sistema"), ptt("Parametro"), valor
                        end if

                        ptt.MoveNext
                    loop until ptt.eof
                end if
            ptt.close: set ptt = nothing

        con.close: set con = nothing            
    end Sub

    Sub ProcesarLista(Sistema, Parametro, CampoDefault)
        dim ptt, con, sqlCommand, sqlTable, valor

        set con = Server.CreateObject("ADODB.Connection")
        con.open Application("Conn")
            con.execute("UPDATE seg_Parametros_Valores SET PorDefecto = 0 WHERE (Sistema = '" & Sistema & "') AND (Parametro = '" & Parametro & "');")
            con.execute("UPDATE seg_Parametros_Valores SET PorDefecto = 1 WHERE (Sistema = '" & Sistema & "') AND (Parametro = '" & Parametro & "') AND (Valor = '" & CampoDefault & "');")
        con.close: set con = nothing
    End Sub

    function limpiar(cadena)
        limpiar = Replace(cadena,"'","´")    
    end function

    sub ActualizarSistema()
        Nuevo = request.form("Nuevo")

        Sistema = request.form("Codigo")
        Nombre = limpiar(request.form("Nombre"))
        Descripcion = limpiar(request.form("Descripcion"))
        ClaseApp = Request.Form("ClaseApp")
        IndiceOrdenamiento = Request.Form("IndiceOrdenamiento")
        Icono = Request.Form("Icono")
        sBitacora = Request.Form("sBitacora")

        if Nuevo = "1" then
            sqlString = "INSERT INTO seg_Sistemas(sysCodigo, sysNombre, sysDescripcion, sysWeb, sysMenuIndice, sysIcon, sysBitacora) " & _
                             "VALUES ('" & Sistema & "', '" & Nombre & "', '" & Descripcion & "', " & ClaseApp & ", " & IndiceOrdenamiento & ", '" & Icono & "', " & sBitacora & ");"
        else
            sqlString = "UPDATE seg_Sistemas " & _
                           "SET sysNombre = '" & Nombre & "'," & _
                              " sysDescripcion = '" & Descripcion & "'," & _
                              " sysWeb = " & ClaseApp & "," & _
                              " sysMenuIndice = " & IndiceOrdenamiento & "," & _
                              " sysIcon = '" & Icono & "'," & _
                              " sysBitacora = " & sBitacora & _
                        " WHERE (sysCodigo = '" & Sistema & "');"
        end if

        set cc = Server.CreateObject("ADODB.Connection")
        cc.open Application("Conn")
            cc.execute(sqlString)    
        cc.close: set cc = nothing 
    end sub


    '
    ' Main()
    '
    
    if Request.Cookies("usuario") = "" then
        Response.Redirect "../login.asp"      
    end if

    ActualizarSistema

    if TieneVariables(Sistema) > 0 then 
        ProcesarVariables(Sistema) 
    end if

   Response.redirect "lista.asp?op=" & Request.Form("ordenadoPor")
%>