<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->
        <% PageTitle = "Aplicar Transacción" %>
        <title><%= PageTitle %></title>

        <%
            dim ccon, pre, editor, llave, sqlString, aplicado, tt
            dim secItem, trAplicado

            llave = request.QueryString("registro")
            trAplicado = EstatusTransaccion(llave)

            ' Funciones y Procedimientos ----------------------------------------------------------------------------------
                function EstatusTransaccion(llave)
                    dim con, tt, sqlString

                    sqlString = "SELECT Aplicado FROM pre_Presupuesto_Detalles WHERE Llave = " & llave & ";"

                    set con = Server.CreateObject("ADODB.Connection")
                    con.open Application("Conn")
                        set tt = con.Execute(sqlString)
                            EstatusTransaccion = tt("Aplicado")
                        tt.close: set tt = nothing
                    con.close: set con = nothing
                end function   

                function PreLlaveItemLista(llave)
                    dim con, tt, sqlString

                    sqlString = "SELECT ISNULL(ItemLista, -1) AS Item FROM pre_Presupuesto_Detalles WHERE Llave = " & Llave & ";"

                    set con = Server.CreateObject("ADODB.Connection")
                    con.open Application("Conn")                    
                        set tt = con.Execute(sqlString)
                            PreLlaveItemLista = tt("Item")
                        tt.close: set tt = nothing
                    con.close: set con = nothing
                end function

                sub BorrarItemLista(Item)
                    dim con, sqlString

                    sqlString = "DELETE FROM pre_Listas_Detalles WHERE Secuencia = " & Item & ";"
                    set con = Server.CreateObject("ADODB.Connection")
                    con.open Application("Conn")                    
                        con.Execute(sqlString)
                    con.close: set con = nothing
                end sub

                function CuentaDestino(llave)
                    dim con, tt, sqlString

                    sqlString = "SELECT CuentaDestino FROM pre_Presupuesto_Detalles WHERE Llave = " & Llave & ";"

                    set con = Server.CreateObject("ADODB.Connection")
                    con.open Application("Conn")
                        set tt = con.Execute(sqlString)
                            CuentaDestino = tt("CuentaDestino")
                        tt.close: set tt = nothing
                    con.close: set con = nothing
                end function

                sub MoverPunteroCuentaCompartida (Llave)
                    dim con, tt, sqlString, usuario, PrimeraSecuencia
                    dim punteroEncontrado, SecuenciaApuntador, Cuenta, sw

                    punteroEncontrado = 0
                    usuario = Request.Cookies("Usuario")
                    cuenta = CuentaDestino(llave)

                    sqlString = "SELECT Secuencia, Contacto, UltimaFechaAplicada, Puntero " & _
                                "FROM dbo.pre_Cuentas_Comparticiones AS co " & _
                                "WHERE (Usuario = '" & usuario & "') " & _
                                "AND (Cuenta = '" & Cuenta & "')"
                    
                    set con = Server.CreateObject("ADODB.Connection")
                    con.open Application("Conn")                                
                        set tt = con.Execute(sqlString)
                            if not (tt.bof or tt.eof) then
                                PrimeraSecuencia = tt("Secuencia")                              ' Es una Cuenta Compartida
                                sw = false                                                      ' Buscamos el puntero...

                                Do
                                    if tt("Puntero") = 1 then
                                        sw = true
                                    else
                                        tt.MoveNext
                                    end if                            
                                Loop Until (tt.eof OR sw)

                                if sw then                                                      ' Encontramos el puntero actual o llegamos a fin de archivo
                                    tt.MoveNext                                                 ' Si encontramos el puntero, nos movemos al siguiente

                                    if tt.eof then                                              ' Si llegamos al fin de archivo, debemos apuntar al primero registro de la lista               
                                        SecuenciaApuntador = PrimeraSecuencia
                                    else
                                        SecuenciaApuntador = tt("Secuencia")
                                    end if
                                else
                                    SecuenciaApuntador = PrimeraSecuencia                       ' Nadie tiene el apuntador... Apuntamos al primer contacto
                                end if
                                                                                                ' Actualizamos el puntero...
                                sqlString = "UPDATE pre_Cuentas_Comparticiones " &  _           
                                            "SET UltimaFechaAplicada = NULL, " & _
                                                " Puntero = 0 " & _
                                            "WHERE (Usuario = '" & usuario & "') " & _
                                            "AND (Cuenta = '" & Cuenta & "');"

                                con.Execute(sqlString)

                                sqlString = "UPDATE pre_Cuentas_Comparticiones " &  _
                                            "SET UltimaFechaAplicada = CAST(LEFT(dbo.sysDateTimeOffset(), 16) AS DateTime), " & _
                                                " Puntero = 1 " & _
                                            "WHERE (Usuario = '" & usuario & "') " & _
                                            "AND (Cuenta = '" & Cuenta & "') " & _
                                            "AND (Secuencia = " & SecuenciaApuntador & ");"

                                con.Execute(sqlString)
                            end if

                        tt.close: set tt = nothing
                    con.close: set con = nothing
                end sub                 
            ' -------------------------------------------------------------------------------------------------------------
        %>        
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

        <div class="page-title-bar">
            <div class="title"><%= PageTitle %></div>
        </div>

        <main>
            <%
                set ccon = Server.CreateObject("ADODB.Connection")
                ccon.open Application("Conn")
                    if trAplicado = 0 then
                        sqlString = "UPDATE pre_Presupuesto_Detalles SET Aplicado = 1 WHERE Llave = " & llave & ";"
                        ccon.Execute(sqlString)

                        '
                        ' Si la linea contiene un Item de alguna lista, será eliminado
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
                    else
                        sqlString = "UPDATE pre_Presupuesto_Detalles SET Aplicado = 0 WHERE Llave = " & llave & ";"
                        ccon.Execute(sqlString)                
                    end if

                    response.redirect "pre_det_editar.asp"                
                ccon.close: set ccon = nothing
            %>
        </main>


        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>