<%@ CodePage=65001 %>
<% Response.Buffer = True %>

<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
  	<head>
    	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
		<!-- #include virtual = "/core/includes/menu_pass.inc" -->

		<%
			Dim conexion, tabla, usuario, password, SQLString, mantener, menubar

			set conexion = Server.CreateObject("ADODB.Connection")
			conexion.open Application("Conn")					

			Usuario = Ucase(Request.Form("txtUsuario"))
			password = UCase(Request.Form("txtPassword"))
			mantener = Request.Form("chkMantener")
			menubar =  Request.Form("chkMenu")		

			Sub PrepararCookies()
				For Each cookie in Request.Cookies
					Response.Cookies(cookie).Expires = DateAdd("d", -1, now())
				Next				

				Response.Cookies("usuario") = usuario
				Response.Cookies("nombre") = tabla("Nombre")
				Response.Cookies("usuPath") = "/perfiles/" & usuario
				Response.Cookies("max_WP") = ContarWallpapers()
				Response.Cookies("usu_WP") = wallPaperUsuario()

				Response.Cookies("usuario").Expires = Date() + 1
				Response.Cookies("nombre").Expires = Date() + 1
				Response.Cookies("usuPath").Expires = Date() + 1
				Response.Cookies("max_WP").Expires = Date() + 1
				Response.Cookies("usu_WP").Expires = Date() + 1
			End Sub			

			Sub NoExpirar()
				Response.Cookies("usuario").Expires = Date() + 3650
				Response.Cookies("nombre").Expires = Date() + 3650
				Response.Cookies("usuPath").Expires = Date() + 3650					
				Response.Cookies("max_WP").Expires = Date() + 3650					
				Response.Cookies("usu_WP").Expires = Date() + 3650				
			End Sub					
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
						if menubar = 1 then 
							CrearMenu usuario
							ActualizarSnippetsUsuario usuario
						end if
						PrepararCookies						
						if mantener = 1 then NoExpirar	

						'Response.Cookies("noVerEncabezado") = 0
						Response.Redirect "/core"
					else
						Response.Redirect "login.asp"
					end if
					
					tabla.Close: Set tabla=nothing
				end if
			else
				Response.Redirect "login.asp"		
			end if

			conexion.close: set conexion=nothing			
		%>	
	</body>
</html>