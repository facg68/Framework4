<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
	dim Con, t, sqlString, Sistema

	Rol = request.querystring("r")
	ordenadoPor = request.querystring("o")
	Sistema = request.querystring("s")

	set Con = Server.CreateObject("ADODB.Connection")
	Con.open Application("Conn")

	'
	' Funciones y Procedimientos
	'

	function SistemaAsignado(Rol)
		dim res, sqlCommand

		sqlCommand = "SELECT ISNULL(CodigoSys, '*') AS Codigo FROM seg_Roles WHERE rolCodigo = '" & Rol & "';"

		set res = Con.execute(sqlCommand)
			if not (res.bof or res.eof) then
				SistemaAsignado = res("Codigo")
			else
				SistemaAsignado = "*"
			end if
		res.close: set res = nothing
	end function

	'
	' Main()
	'

	if SistemaAsignado(Rol) = "*" then
		sqlString = "SELECT COUNT(*) AS Cuantos FROM seg_Sistemas;"	

		set t = Con.execute(sqlString)   
			if t("Cuantos") > 5 then
				response.redirect "editar_rol_tabs.asp?r=" & Rol & "&o=" & ordenadoPor & "&s=" & Sistema
			else
				response.redirect "editar_rol_tabla.asp?r=" & Rol & "&o=" & ordenadoPor & "&s=" & Sistema
			end if
		t.close: set t = nothing
	else
		response.redirect "editar_rol_unico.asp?r=" & Rol & "&o=" & ordenadoPor & "&s=" & Sistema
	end if
	
	con.close: set con = nothing
%>