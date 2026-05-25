<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html lang="es">
    <head>
        <!-- #include virtual = "/forma/mobile/recursos/includes/header.inc" -->     
        <% PageTitle = "Borrar Presupuesto" %>
        <title><%= PageTitle %></title>

		<%
			function TieneTransacciones(Usuario, Presupuesto)
				dim cc, tt, sqlString
		
				sqlString = "SELECT ISNULL(COUNT(*), 0) AS Cuantos " & _
							  "FROM pre_Presupuesto_Detalles " & _
							 "WHERE (Usuario = '" & Usuario & "') " & _
							   "AND (Presupuesto = '" & Presupuesto & "') " & _
							   "AND (Aplicado = 1);"
				
				set cc = Server.CreateObject("ADODB.Connection")
				cc.open Application("Conn")
					set tt = cc.Execute(sqlString)
						TieneTransacciones = tt("Cuantos")				
					tt.close: set tt = nothing
				cc.close: set cc = nothing
			end function
		%>		        
    </head>

    <body>
        <!-- #include virtual = "/forma/mobile/recursos/includes/menu.inc" -->     

		<%
			dim usu, pre, c, sqlString 

			usu = Request.Cookies("Usuario")
			pre = Request.QueryString("pre")
        %>

        <div class="page-title-bar">
            <%= PageTitle %>
        </div>

        <main>
            <div class="contenedor">
                <%
                    if TieneTransacciones(usu, pre) = 0 then   
                        set c = server.CreateObject("ADODB.Connection")
                        c.open Application("Conn")
                            '
                            '  Borramos los detalles
                            '
                            sqlString = "DELETE FROM pre_Presupuesto_Detalles " & _
                                        "WHERE Usuario = '" & usu & "' " & _
                                            "AND Presupuesto = '" & pre & "';"

                            c.execute (sqlString)

                            '
                            '  Borramos el encabezado
                            '
                            sqlString = "DELETE FROM pre_Presupuesto_Encabezado " & _
                                        "WHERE Usuario = '" & usu & "' " & _
                                            "AND Presupuesto = '" & pre & "';"

                            c.execute (sqlString)
                        c.close: set c = nothing
                        
                        response.redirect "../lista.asp"
                    end if  
                %>

                Error:

                <br /><br />

                Este presupuesto no se puede borrar porque ya tiene (por lo menos) 
                una transaccion aplicada que afecta las cuentas.

                <br /><br />

                Si no se han cerrado las cuentas (en el módulo de cuentas), 
                aún puede editar las transacciones para quitar el estatus de 
                "aplicado", de esa forma podrá eliminar éste presupuesto.

                <br /><br />

                Fin de Proceso            
            <div>
        </main>

        <footer class="footer-contextual">
            <button class="footer-button" type="button" aria-label="Volver" onclick="volver()">
                <i class="fas fa-undo-alt"></i>
            </button>
        </footer>

        <script>
            function volver() {
                var vinculo = "../lista.asp";
                window.location.href = vinculo;
            }
        </script>

        <!-- #include virtual = "/forma/mobile/recursos/includes/close.inc" -->     
    </body>
</html>