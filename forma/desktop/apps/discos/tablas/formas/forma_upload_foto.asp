<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim p, separador, cc, vinculo, fOriginal, fNuevo
    dim tam, donde, fso, puntoSeparador, ext

    p = lcase(Request.Cookies("usuPath")) & "/discos"

    '
    ' Subimos la foto a la carpeta...
    '

    Set Upload = Server.CreateObject("Persits.Upload") 
    count = upload.savevirtual(p)

    response.write "Files:<BR>"

    For Each File in Upload.Files
        '
        ' Buscamos el nombre del archivo
        ' Aunque esta dentro de un Loop, solo hay
        ' un solo item            
        '
        fOriginal = cStr(File.Path)
    Next

    '
    ' Separamos el Path del nombre del archivo original           
    '

    tam = len(fOriginal)
    donde = -1
    separador = 0            

    for k = tam to 1 step -1
        if mid(fOriginal, k, 1) = "\" then
            if (donde = -1) then
                donde = k
            end if
        end if
    next

    if donde <> -1 then
        separador = donde
    end if

    '
    ' buscamos el tipo de archivo...
    '

    tam = len(fOriginal)
    donde = -1
    puntoSeparador = 0            

    for k = tam to 1 step -1
        if mid(fOriginal, k, 1) = "." then
            if (donde = -1) then
                donde = k
            end if
        end if
    next

    if donde <> -1 then
        puntoSeparador = donde
        ext = right(fOriginal, (tam-donde))
    else
        ext = "gif"
    end if

    '
    ' Cargamos los parámetros          
    '

    For Each Item in Upload.Form
        select case Item.Name
            case "Forma": Forma = Item.Value
        end select
    Next

    '
    ' Copiamos el archivo original con el nombre
    ' que le corresponde y luego borramos el original
    '
    fObjeto = Forma & "." & ext
    fNuevo = left(fOriginal, separador) & Forma & "." & ext

    Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFile fOriginal, fNuevo
        fso.DeleteFile fOriginal
    set fso = nothing    

    '
    ' Acualizamos el Registro...
    '

    sqlString = "UPDATE dbo.discos_Formas " & _
                "SET Icono_Forma = '" & fObjeto & "' " & _
                "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                "AND (Forma = '" & Forma & "');"
    
    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")
        cc.execute(sqlString)
    cc.close: set cc = nothing
    
    '
    ' Volvemos al formulario de edicion
    '

    vinculo = "lista.asp"

    response.redirect vinculo
%>