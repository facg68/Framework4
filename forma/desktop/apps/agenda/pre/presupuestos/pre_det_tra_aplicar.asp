<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function PreLlaveItemLista(llave)
                dim cc, tt, sqlString

                sqlString = "SELECT ISNULL(ItemLista, -1) AS Item FROM pre_Presupuesto_Detalles WHERE Llave = " & Llave & ";"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.Execute(sqlString)

                PreLlaveItemLista = tt("Item")

                tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function

            sub BorrarItemLista(Item)
                dim cc, sqlString

                sqlString = "DELETE FROM pre_Listas_Detalles WHERE Secuencia = " & Item & ";"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                    cc.Execute(sqlString)
                cc.close: set cc = nothing            
            end sub

            function CuentaDestino(llave)
                dim cc, tt, sqlString

                sqlString = "SELECT CuentaDestino FROM pre_Presupuesto_Detalles WHERE Llave = " & Llave & ";"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.Execute(sqlString)

                CuentaDestino = tt("CuentaDestino")

                tt.close: set tt = nothing
                cc.close: set cc = nothing            
            end function

            sub MoverPunteroCuentaCompartida (Llave)
                dim cc, tt, sqlString, usuario, PrimeraSecuencia
                dim punteroEncontrado, SecuenciaApuntador, Cuenta, sw

                punteroEncontrado = 0
                usuario = Request.Cookies("Usuario")
                cuenta = CuentaDestino(llave)

                sqlString = "SELECT Secuencia, Contacto, UltimaFechaAplicada, Puntero " & _
                              "FROM dbo.pre_Cuentas_Comparticiones AS co " & _
                             "WHERE (Usuario = '" & usuario & "') " & _
                               "AND (Cuenta = '" & Cuenta & "')"
                
                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                set tt = cc.Execute(sqlString)
                    if not (tt.bof or tt.eof) then
                        '
                        ' Es una Cuenta Compartida
                        '
                        PrimeraSecuencia = tt("Secuencia")

                        '
                        ' Buscamos el puntero...
                        '
                        sw = false

                        Do
                            if tt("Puntero") = 1 then
                                sw = true
                            else
                                tt.MoveNext
                            end if                            
                        Loop Until (tt.eof OR sw)

                        '
                        ' Aqui hay dos opciones... O encontramos el puntero actual,
                        ' o llegamos a fin de archivo sin encontrar el apuntador
                        '

                        if sw then
                            '
                            ' Si encontramos el puntero, nos movemos al siguiente
                            ' registro, que debería ser el nuevo apuntador
                            '
                            tt.MoveNext

                            '
                            ' Pero si llegamos al fin de archivo, debemos 
                            ' apuntar al primero registro de la lista 
                            ' (PrimeraSecuencia)
                            '

                            if tt.eof then
                                SecuenciaApuntador = PrimeraSecuencia
                            else
                                SecuenciaApuntador = tt("Secuencia")
                            end if
                        else
                            '
                            ' Nadie tiene el apuntador... Empezamos desde
                            ' el principio, apuntando al primer contacto
                            ' de la lista (PrimeraSecuencia)
                            '
                            SecuenciaApuntador = PrimeraSecuencia
                        end if

                        '
                        ' Actualizamos el puntero...
                        '

                        sqlString = "UPDATE pre_Cuentas_Comparticiones " &  _
                                       "SET UltimaFechaAplicada = NULL, " & _
                                          " Puntero = 0 " & _
                                     "WHERE (Usuario = '" & usuario & "') " & _
                                       "AND (Cuenta = '" & Cuenta & "');"

                        cc.Execute(sqlString)

                        sqlString = "UPDATE pre_Cuentas_Comparticiones " &  _
                                       "SET UltimaFechaAplicada = CAST(LEFT(dbo.sysDateTimeOffset(), 16) AS DateTime), " & _
                                          " Puntero = 1 " & _
                                     "WHERE (Usuario = '" & usuario & "') " & _
                                       "AND (Cuenta = '" & Cuenta & "') " & _
                                       "AND (Secuencia = " & SecuenciaApuntador & ");"

                        cc.Execute(sqlString)
                    end if

                tt.close: set tt = nothing
                cc.close: set cc = nothing
            end sub
        %>
    </head>

    <body>
        <%
            dim con, pre, editor, llave, sqlString, aplicado, tt
            dim secItem

            pre = request.QueryString("p")
            llave = request.QueryString("l")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            set tt = con.Execute("SELECT Aplicado FROM pre_Presupuesto_Detalles WHERE Llave = " & llave & ";")
                aplicado = tt("Aplicado")
            tt.close: set tt = nothing

            if aplicado = 0 then
                sqlString = "UPDATE pre_Presupuesto_Detalles SET Aplicado = 1 WHERE Llave = " & llave & ";"
                con.Execute(sqlString)
                con.close: set con = nothing

                '
                ' Si la linea conetiene un Item de alguna lista, será eliminado
                ' de la lista y de la base de datos sin posibilidad a recuperarlo
                ' nuevamente
                '
                secItem = PreLlaveItemLista(llave)

                if secItem <> "-1" then
                    BorrarItemLista secItem
                end if

                '
                ' Si es una cuenta compartida, verificamos el puntero del contacto
                '
                MoverPunteroCuentaCompartida Llave
            end if

            vinculo = "pre_det_editar.asp?p=" & pre & "&d=" & request.QueryString("d") & "&v=" & request.QueryString("v") & "&t=" & request.QueryString("t") & "&e=" & request.QueryString("e") & "&o=" & request.QueryString("o")
            response.redirect vinculo            
        %>    
    </body>
</html>