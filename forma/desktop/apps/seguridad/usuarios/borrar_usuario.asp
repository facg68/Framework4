<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    '
    ' Init()
    '

    dim cc, tt, sqlString, usuario, ordenadoPor

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")

    '
    ' Funciones
    '

    Sub BorrarFolder(Usuario)
        On Error Resume Next

        map = Server.MapPath("\perfiles\")
        folder = map & "\" & Usuario

        Set fso = CreateObject("Scripting.FileSystemObject")
            If fso.FolderExists(folder) Then
                fso.DeleteFolder folder, True  ' True = forzar borrado (incluye contenido)
                
                If Err.Number <> 0 Then
                    ' Manejo del error
                    Response.Write "Error al borrar carpeta: " & Err.Description
                    Err.Clear
                End If
            Else
                ' Opcional: manejar si no existe
                ' Response.Write "La carpeta no existe"
            End If
        Set fso = Nothing

        On Error GoTo 0
    End Sub

    Sub BorrarDatosPrimordiales(Usuario)
        cc.execute("DELETE FROM discos_Carpetas             WHERE (Usuario = '" & Usuario & "');")
        cc.execute("DELETE FROM discos_Tiendas              WHERE (Usuario = '" & Usuario & "');")
        cc.execute("DELETE FROM discos_Casas                WHERE (Usuario = '" & Usuario & "');")
        cc.execute("DELETE FROM discos_formas               WHERE (Usuario = '" & Usuario & "');")
        cc.execute("DELETE FROM discos_Tipos                WHERE (Usuario = '" & Usuario & "');")
        cc.execute("DELETE FROM discos_Plataformas          WHERE (Usuario = '" & Usuario & "');")
        cc.execute("DELETE FROM discos_Clasificaciones      WHERE (Usuario = '" & Usuario & "');")
        cc.execute("DELETE FROM discos_FormatosPantalla     WHERE (Usuario = '" & Usuario & "');")
        cc.execute("DELETE FROM discos_Graficas_Defaults    WHERE (Usuario = '" & Usuario & "');")
        cc.execute("DELETE FROM discos_Idiomas              WHERE (Usuario = '" & Usuario & "');")
        cc.execute("DELETE FROM cal_Calendarios             WHERE (Usuario = '" & Usuario & "');")
        cc.execute("DELETE FROM con_Contactos_Tipos         WHERE (Usuario = '" & Usuario & "');")
        cc.execute("DELETE FROM con_Contactos_Categorias    WHERE (Usuario = '" & Usuario & "');")
        cc.execute("DELETE FROM dbo.pre_Cuentas             WHERE (Usuario = '" & Usuario & "');")
    End Sub
    

    '
    ' Main
    '

    codigo = Request.QueryString("u")   
    ordenadoPor = Request.QueryString("o")   

    sqlString = "DELETE FROM seg_Usuarios WHERE usuCodigo = '" & codigo & "';"
    cc.execute(sqlString)    

    '--------------------------------------'
    ' Proceso para la Extranet de Fabrizio '
    '--------------------------------------'
        BorrarDatosPrimordiales codigo
        BorrarFolder codigo        
    '-----'
    ' Fin '
    '-----'


    cc.close: set cc = nothing

    Response.redirect "lista.asp?o=" & ordenadoPor
%>