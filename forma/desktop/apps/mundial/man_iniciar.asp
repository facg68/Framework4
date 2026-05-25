<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
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
        %>    
    </head>

    <body>
        <%
            dim key, cc

			if (Usuario_Valido() = 1) then
                key = Request.Cookies("mundial_key")

                if key = "dslkwrtywbfsjbvwegowuienweixmlkjri" then
                    set cc = Server.CreateObject("ADODB.Connection")
                    cc.open Application("Conn")

                    '
                    ' Borramos el Historial de Partidos
                    '
                    cc.execute("DELETE FROM mundial_Resultados;")

                    '
                    ' Borramos TODAS LAS POLLAS en el sistema...
                    '
                    cc.execute("DELETE FROM mundial_Apuestas_Det;")
                    cc.execute("DELETE FROM mundial_Apuestas_Enc;")

                    '
                    ' Iniciamos el Master...
                    '
                    cc.execute("UPDATE mundial_Master SET CodigoEquipo = '-';")

                    '
                    ' Activamos TODOS los Equipos que tengan un GRUPO asignado...
                    '
                    cc.execute("UPDATE mundial_Tabla_Equipos SET Activo = 0, EnJuego = 0;")
                    cc.execute("UPDATE mundial_Tabla_Equipos SET Activo = 1, EnJuego = 1 WHERE GRUPO <> '';")

                    '
                    ' Iniciamos los Estatus de Activacion y Creacion de Pollas
                    '
                    cc.execute("UPDATE mundial_Estatus SET Estatus = 1 WHERE Codigo = 'Activar';")
                    cc.execute("UPDATE mundial_Estatus SET Estatus = 0 WHERE Codigo = 'Finalizada';")
                    cc.execute("UPDATE mundial_Estatus SET Estatus = 1 WHERE Codigo = 'Polla';")
                    cc.execute("UPDATE mundial_Estatus SET Estatus = 5 WHERE Codigo = 'Boleto';")
                    cc.execute("UPDATE mundial_Estatus SET Estatus = 1 WHERE Codigo = 'Etapa';")

                    '
                    ' Listo! El Sistema ha sido reiniciado y está listo para empezar!!!
                    '
                    cc.close: set cc = nothing
                end if
            end if

            response.redirect "mundial.asp"
        %>
    </body>
</html>
