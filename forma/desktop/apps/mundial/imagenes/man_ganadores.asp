<html>
    <head>
        <%
            function EquipoGanador()
                dim con, t, sqlString

                sqlString = "SELECT CodigoEquipo " & _
                            "FROM mundial_Master " & _
                            "WHERE IndiceCuadro = 31;"
                
                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")
                set t = con.execute(sqlString)

                EquipoGanador = t("CodigoEquipo")

                t.close: set t = nothing
                con.close: set con = nothing
            end function

            function CuantosAtinaronAlGanador()
                dim con, t, sqlString

                sqlString = "SELECT COUNT(*) AS Cuantos " & _
                            "FROM mundial_Totales_Globales " & _
                            "WHERE (Estatus = 1) " & _
                            "AND (Ganador = '" & EquipoGanador() & "');"    
                
                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")
                set t = con.execute(sqlString)

                CuantosAtinaronAlGanador = t("Cuantos")

                t.close: set t = nothing
                con.close: set con = nothing
            end function

            function MaximoPuntajeGanador()
                dim con, t, sqlString

                sqlString = "SELECT MAX(Puntaje) AS Maximo " & _
                            "FROM dbo.mundial_Totales_Globales AS t " & _
                            "WHERE (Ganador = '" & EquipoGanador() & "') " & _
                            "AND (Estatus = 1);"
                
                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")
                set t = con.execute(sqlString)

                MaximoPuntajeGanador = t("Maximo")

                t.close: set t = nothing
                con.close: set con = nothing                
            end function

            function MaximoPuntaje()
                dim con, t, sqlString

                sqlString = "SELECT MAX(Puntaje) AS Maximo " & _
                            "FROM dbo.mundial_Totales_Globales AS t " & _
                            "WHERE (Estatus = 1);"
                
                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")
                set t = con.execute(sqlString)

                MaximoPuntaje = t("Maximo")

                t.close: set t = nothing
                con.close: set con = nothing                
            end function            

            function UnicoGanador()
                dim con, t, sqlString

                sqlString = "SELECT Secuencia " & _
                            "FROM dbo.mundial_Totales_Globales AS t " & _
                            "WHERE (Ganador = '" & EquipoGanador() & "') " & _
                            "AND (Estatus = 1) ;"
                
                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")
                set t = con.execute(sqlString)

                UnicoGanador = t("Secuencia")

                t.close: set t = nothing
                con.close: set con = nothing                            
            end function

			Function Usuario_Valido()
				dim con, t, sqlString
				
				Usuario_Valido = 0
				sqlString = "exec seg_pa_VerificarPermisoUsuario '" & Request.Cookies("Usuario") & "', 'mundial', 'mundial.050'"

				set con = Server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)

				if t("Acceso") = 1 then
					Usuario_Valido = 1
				end if
				
				t.close: set t=nothing
				con.close: set con=nothing			
			End Function	 

            Sub HistorialCompleto()
            	dim con

                set con = Server.CreateObject("ADODB.Connection")
                con.open Application("Conn")
                con.execute("UPDATE mundial_Estatus SET Estatus = 6 WHERE Codigo = 'Etapa';")
                con.close: set con = nothing		            
            End Sub           
        %>    
    </head>

    <body>
        <%
            dim con, t, sqlString, key
            dim AlGanador, maxGanador

			if (Usuario_Valido() = 1) then
                key = Request.Cookies("mundial_key")

                if key = "dslkwrtywbfsjbvwegowuienweixmlkjri" then

                    set con = Server.CreateObject("ADODB.Connection")
                    con.open Application("Conn")

                        '
                        ' Llegó el momento de la verdad!
                        '
                        ' Actualizamos todas las apuestas para que
                        ' el Campo "Ganador" sea igual a CERO
                        '
                        con.execute("UPDATE dbo.mundial_Apuestas_Enc SET Ganador = 0;")

                        '
                        ' Verifiquemos Cuantas Personas atinaron al 
                        ' vencedor de la copa mundial...
                        '
                        AlGanador = CuantosAtinaronAlGanador()

                        if AlGanador = 0 then
                            '
                            ' Ok, tenemos que verificar puntos...
                            '
                            maxGanador = MaximoPuntaje()

                            set t = con.execute("SELECT Secuencia " & _
                                                "FROM dbo.mundial_Totales_Globales " & _
                                                "WHERE (Estatus = 1) " & _
                                                "AND (Puntaje = " & maxGanador & ");")
                            
                                Do
                                    sqlString = "UPDATE mundial_Apuestas_Enc " & _
                                                "SET Ganador = 1 " & _
                                                "WHERE Secuencia = '" & t("Secuencia") & "';"

                                    con.execute(sqlString)
                                    t.movenext
                                Loop Until t.eof

                            t.close: set t = nothing

                        else
                            '
                            ' UNO O MAS jugadores atinaron al Equipo Ganador!
                            ' Veamos quien se lleva el acumulado!
                            '

                            if AlGanador = 1 then
                                '
                                ' Bien! Tenemos un UNICO GANADOR!!!
                                '

                                sqlString = "UPDATE mundial_Apuestas_Enc " & _
                                            "SET Ganador = 1 " & _
                                            "WHERE Secuencia = '" & UnicoGanador() & "';"

                                con.execute(sqlString)
                            else
                                '
                                ' Tenemos que verificar puntajes...
                                '
                                maxGanador = MaximoPuntajeGanador()

                                set t = con.execute("SELECT Secuencia " & _
                                                    "FROM dbo.mundial_Totales_Globales " & _
                                                    "WHERE (Estatus = 1) " & _
                                                    "AND (Ganador = '" & EquipoGanador() & "') " & _
                                                    "AND (Puntaje = " & maxGanador & ");")
                                
                                    Do
                                        sqlString = "UPDATE mundial_Apuestas_Enc " & _
                                                    "SET Ganador = 1 " & _
                                                    "WHERE Secuencia = '" & t("Secuencia") & "';"

                                        con.execute(sqlString)
                                        t.movenext
                                    Loop Until t.eof

                                t.close: set t = nothing

                            end if
                        end if

                        '
                        ' Ya tenemos al Ganador (o ganadores!)
                        ' Cerramos, oficialmente, la competencia
                        '

                        sqlString = "UPDATE mundial_Estatus SET Estatus = 1 WHERE Codigo = 'Finalizada';"
                        con.execute(sqlString)

                    con.close: set con = nothing

                    HistorialCompleto
                end if
            end if

            response.redirect "mundial.asp"
        %>
    </body>
</html>
