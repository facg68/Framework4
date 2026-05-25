<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
  	<head>
    	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
		<!-- #include virtual = "/core/includes/menu_pass.inc" -->

		<%
			Dim conexion, usuIndice, usuMenu, tabla, usuario, password, SQLString, mantener, menubar, tt

			set conexion = Server.CreateObject("ADODB.Connection")
			conexion.open Application("Conn")					

			Usuario = Ucase(Request.Form("txtUsuario"))
			password = UCase(Request.Form("txtPassword"))
			mantener = Request.Form("chkMantener")
			menubar =  Request.Form("chkMenu")	
		%>		
	</head>
	
	<body>
		<%
			if UsuarioExiste(usuario) = 1 then
				if ResetearPassword(usuario) = 1 then
					Session("c2_Reset") = "/*" & CadenaRandom() & "$|"
					Response.cookies("c2_Reset") = Session("c2_Reset")
					Response.Cookies("rUsuario") = usuario					

					response.redirect "pass.asp"
				else
					SQLString = "seg_pa_Usuarios '" & usuario & "', '" & password & "'"
			
					set tabla = conexion.execute(SQLString)
					
					if NOT (tabla.bof or tabla.eof) then
						Response.Cookies("usuario") = usuario
						Response.Cookies("nombre") = tabla("Nombre")
						Response.Cookies("usuPath") = "/perfiles/" & usuario
						Response.Cookies("max_WP") = ContarWallpapers()
						Response.Cookies("usu_WP") = wallPaperUsuario()
						

						'
						' El Servidor Ha Aceptado al usuario...
						' Creamos su WebMenu
						'
						if menubar = 1 then
							CrearMenu Request.Cookies("usuario")
						end if

						if mantener = 1 then
							'
							' La conexión expira cuando el usuario
							' decida salir de la app
							'
							Response.Cookies("usuario").Expires = Date() + 365
							Response.Cookies("nombre").Expires = Date() + 365
							Response.Cookies("usuPath").Expires = Date() + 365					
							Response.Cookies("max_WP").Expires = Date() + 365					
							Response.Cookies("usu_WP").Expires = Date() + 365					
						else
							'
							' La conexión expira en 24 horas
							'
							Response.Cookies("usuario").Expires = Date() + 7
							Response.Cookies("nombre").Expires = Date() + 7
							Response.Cookies("usuPath").Expires = Date() + 7
							Response.Cookies("max_WP").Expires = Date() + 7
							Response.Cookies("usu_WP").Expires = Date() + 7
						end if

						'Response.Cookies("noVerEncabezado") = 0
						Response.Redirect "/forma/mobile/"
					else
						Response.Redirect "/forma/mobile/login.asp"
					end if
					
					tabla.Close: Set tabla=nothing
				end if
			else
				Response.Redirect "/forma/mobile"		
			end if

			conexion.close: set conexion=nothing	
		%>	
	</body>
</html>