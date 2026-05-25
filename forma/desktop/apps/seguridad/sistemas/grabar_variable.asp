<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <%
        dim cc, sqlString, Nuevo, Sistema, Variable, ConsultaSQL, Tipo, ValorDefault, Descripcion, Afectacion, Consulta

        set cc = Server.CreateObject("ADODB.Connection")
        cc.open Application("Conn")    
    %>

    <head>
        <%
            function limpiar(cadena)
                limpiar = Replace(cadena,"'","´")    
            end function

            Sub BorrarLista(Sistema, Parametro)
                sqlString = "DELETE FROM seg_Parametros_Valores " & _
                             "WHERE Sistema = '" & Sistema & "' " & _
                               "AND Parametro = '" & Parametro & "';"

                cc.execute(sqlString)             
            End Sub

            sub ActualizarLista(Sistema, Parametro)
                vDefault = "/*-!"

                for k = 1 to 15
                    formValor = Request.Form("lValor_" & k)
                    formDescripcion = Request.Form("lDescripcion_" & k)
                    formPorDefecto = Request.Form("lPorDefecto_" & k)

                    if formPorDefecto = 1 then
                        if vDefault = "/*-!" then
                            vDefault = formValor
                        end if
                    else
                        formPorDefecto = 0
                    end if

                    if (len(trim(formValor)) > 0) AND (len(trim(formDescripcion))) then
                        '
                        ' Insertamos valor en la tabla
                        '
                        sqlString = "INSERT INTO seg_Parametros_Valores(Sistema, Parametro, Valor, Descripcion, PorDefecto) " & _
                                         "VALUES ('" & Sistema & "', '" & Parametro & "', '" & formValor & "', '" & formDescripcion & "', " & formPorDefecto & ");"

                        cc.execute(sqlString)
                    end if
                next

                '
                ' Grabamos el Valor por Defecto en la Variable
                '
                if vDefault <> "/*-!" then
                    sqlString = "UPDATE seg_Parametros " & _
                                   "SET ValorDefault = '" & vDefault & "' " & _
                                 "WHERE Sistema = '" & Sistema & "' " & _
                                   "AND Parametro = '" & Parametro & "';"

                    cc.execute(sqlString)
                end if            
            end sub

            function ASCIIstring(CadenaASCII)
                Dim Consulta, Cuantos

                Consulta = Split(CadenaASCII, ",")

                If IsArray(Consulta) Then
                    Cuantos = UBound(Consulta)

                    if Cuantos > 0 then
                        ASCIIstring = ""

                        for k = 0 to Cuantos
                            ASCIIstring = ASCIIstring & CHR(Consulta(k))
                        next
                    end if
                End If
            end function            
        %>
    </head>

    <body>
        <%
            Sistema = request.form("Sistema")
            Nuevo = request.form("Nuevo")

            Variable = Request.Form("Parametro")
            Tipo = Request.Form("TipoParametro")
            Exponer = Request.Form("Exponer")
            Descripcion = Request.Form("Descripcion")
            Afectacion = Request.Form("Afectacion")

            BorrarLista Sistema, Variable           

            if Nuevo = "1" then
                Select Case Tipo
                    case 2
                        '
                        ' Es una Variable, por lo tanto, grabamos su "Valor por Deecto"
                        '
                        ValorDefault = Request.Form("ValorDefault")

                        sqlString = "INSERT INTO seg_Parametros(Sistema, Parametro, TipoParametro, ValorDefault, Descripcion, Afectacion, Exponer) " & _
                                            "VALUES ('" & Sistema & "', '" & Variable & "', " & Tipo & ", '" & ValorDefault & "', '" & Descripcion & "', '" & Afectacion & "', " & Exponer & ");"            

                        cc.execute(sqlString) 

                    case 4
                        '
                        ' Es una lista... Se graba la variable por partes..
                        ' Grabamos la variable sin valor default
                        '
                        sqlString = "INSERT INTO seg_Parametros(Sistema, Parametro, TipoParametro, Descripcion, Afectacion, Exponer) " & _
                                            "VALUES ('" & Sistema & "', '" & Variable & "', " & Tipo & ", '" & Descripcion & "', '" & Afectacion & "', " & Exponer & ");"

                        cc.execute(sqlString) 
                        ActualizarLista Sistema, Variable

                    case 5
                        '
                        ' Es un Query de SQL Server...
                        ' Grabamos la consulta en el valor default
                        '
                        Consulta = ASCIIstring(request.form("ValorConsulta"))

                        sqlString = "INSERT INTO seg_Parametros(Sistema, Parametro, TipoParametro, ConsultaSQL, Descripcion, Afectacion, Exponer) " & _
                                            "VALUES ('" & Sistema & "', '" & Variable & "', " & Tipo & ", '" & Consulta & "', '" & Descripcion & "', '" & Afectacion & "', " & Exponer & ");"            

                        cc.execute(sqlString) 

                    case 6
                        '
                        ' Es un Selector de Colores
                        '
                        ValorDefault = Request.Form("ColorPredeterminado")

                        sqlString = "INSERT INTO seg_Parametros(Sistema, Parametro, TipoParametro, ValorDefault, Descripcion, Afectacion, Exponer) " & _
                                            "VALUES ('" & Sistema & "', '" & Variable & "', " & Tipo & ", '" & ValorDefault & "', '" & Descripcion & "', '" & Afectacion & "', " & Exponer & ");"            

                        cc.execute(sqlString)     

                    case 7
                        '
                        ' Es una Barra de Desplazamiento
                        '
                        ValorDefault = Request.Form("barraDesplazamiento")

                        sqlString = "INSERT INTO seg_Parametros(Sistema, Parametro, TipoParametro, ValorDefault, Descripcion, Afectacion, Exponer) " & _
                                            "VALUES ('" & Sistema & "', '" & Variable & "', " & Tipo & ", '" & ValorDefault & "', '" & Descripcion & "', '" & Afectacion & "', " & Exponer & ");"            

                        cc.execute(sqlString)                                                

                    case else
                        '
                        ' Es un Permiso o un campo "Si / No", por lo que no grabamos un Default
                        '
                        sqlString = "INSERT INTO seg_Parametros(Sistema, Parametro, TipoParametro, Descripcion, Afectacion, Exponer) " & _
                                            "VALUES ('" & Sistema & "', '" & Variable & "', " & Tipo & ", '" & Descripcion & "', '" & Afectacion & "', " & Exponer & ");"            

                        cc.execute(sqlString) 
                        
                end select
            else
                select case Tipo
                    case 2
                        ValorDefault = Request.Form("ValorDefault")

                        sqlString = "UPDATE seg_Parametros " & _
                                    "SET TipoParametro = " & Tipo & ", " & _ 
                                        " ValorDefault = '" & ValorDefault & "'," & _
                                        " Descripcion = '" & Descripcion & "'," & _
                                        " Exponer = " & Exponer & ","  & _                                        
                                        " Afectacion = '" & Afectacion & "' " & _
                                    "WHERE (Sistema = '" & Sistema & "') " & _
                                    "AND (Parametro = '" & Variable & "');"   

                        cc.execute(sqlString)   

                    case 4
                        '
                        ' Es una lista... Se graba la variable por partes..
                        ' Actualizamos la variable sin valor default
                        '
                        sqlString = "UPDATE seg_Parametros " & _
                                    "SET TipoParametro = " & Tipo & ", " & _ 
                                        " ValorDefault = '" & ValorDefault & "'," & _
                                        " Descripcion = '" & Descripcion & "'," & _
                                        " Exponer = " & Exponer & ","  & _                                        
                                        " Afectacion = '" & Afectacion & "' " & _
                                    "WHERE (Sistema = '" & Sistema & "') " & _
                                    "AND (Parametro = '" & Variable & "');"

                        cc.execute(sqlString) 
                        ActualizarLista Sistema, Variable

                    case 5
                        '
                        ' Es una Consulta... Se graba en el valor ConsultaSQL
                        '
                        Consulta = ASCIIstring(request.form("ValorConsulta"))
                                                
                        sqlString = "UPDATE seg_Parametros " & _
                                    "SET TipoParametro = " & Tipo & ", " & _ 
                                        " ConsultaSQL = '" & Consulta & "'," & _
                                        " Descripcion = '" & Descripcion & "'," & _
                                        " Exponer = " & Exponer & ","  & _
                                        " Afectacion = '" & Afectacion & "' " & _
                                    "WHERE (Sistema = '" & Sistema & "') " & _
                                    "AND (Parametro = '" & Variable & "');"

                        cc.execute(sqlString) 

                    case 6
                        ValorDefault = Request.Form("ColorPredeterminado")

                        sqlString = "UPDATE seg_Parametros " & _
                                    "SET TipoParametro = " & Tipo & ", " & _ 
                                        " ValorDefault = '" & ValorDefault & "'," & _
                                        " Descripcion = '" & Descripcion & "'," & _
                                        " Exponer = " & Exponer & ","  & _                                        
                                        " Afectacion = '" & Afectacion & "' " & _
                                    "WHERE (Sistema = '" & Sistema & "') " & _
                                    "AND (Parametro = '" & Variable & "');"   

                        cc.execute(sqlString)    

                    case 7
                        ValorDefault = Request.Form("barraDesplazamiento")

                        sqlString = "UPDATE seg_Parametros " & _
                                    "SET TipoParametro = " & Tipo & ", " & _ 
                                        " ValorDefault = '" & ValorDefault & "'," & _
                                        " Descripcion = '" & Descripcion & "'," & _
                                        " Exponer = " & Exponer & ","  & _                                        
                                        " Afectacion = '" & Afectacion & "' " & _
                                    "WHERE (Sistema = '" & Sistema & "') " & _
                                    "AND (Parametro = '" & Variable & "');"   

                        cc.execute(sqlString)                                                    

                    case else

                        sqlString = "UPDATE seg_Parametros " & _
                                    "SET TipoParametro = " & Tipo & ", " & _ 
                                        " ValorDefault = '1'," & _
                                        " Descripcion = '" & Descripcion & "'," & _
                                        " Exponer = " & Exponer & ","  & _
                                        " Afectacion = '" & Afectacion & "' " & _
                                    "WHERE (Sistema = '" & Sistema & "') " & _
                                    "AND (Parametro = '" & Variable & "');"

                        cc.execute(sqlString) 

                end select
            end if
        %>    
    </body>

    <%
        cc.close: set cc = nothing
        Response.redirect "variables.asp?s=" & sistema & "&o=" & Request.Form("ordenadoPor")
    %>    
</html> 