<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
	'
	' Finalmente, grabamos los datos en el servidor...
	'
	Dim Tiquete, k, campo, con, sqlString, t

	set con = server.CreateObject("ADODB.Connection")
	con.open Application("Conn")
	
	'
	' Primero, Necesitamos ASEGURAR una secuencia...
	'
	
	sqlString = "SELECT dbo.mundial_CodigoUnico() AS T;"
	set t = con.execute(sqlString)
	Tiquete = t("T")
	t.close: set t=nothing
	
	'
	' Ya tenemos un tiquete asegurado y la 
	' secuencia ha sido incrementada
	'
	' Procedemos a Grabar las tablas...
	'
	
	sqlString = "INSERT INTO mundial_Apuestas_Enc(Secuencia, Nombre, Cedula, Telefono, Departamento) " & _
				     "VALUES('" & Tiquete & "','" & request.Form("txtNombre") & "','" & request.Form("txtCedula") & "','" & request.Form("txtTelefono") & "','" & request.Form("txtDepartamento") & "');"

	con.execute(sqlString)
	
	'
	' Y los datos de la apuesta...
	'
	
	for k = 1 to 31
		campo = "Polla" & k
	
		sqlString = "INSERT INTO mundial_Apuestas_Det(Secuencia, Posicion, CodEquipo) " & _
						 "VALUES('" & Tiquete & "'," & k & ",'" & Request.Form(campo) & "');"
		con.execute(sqlString)						 
	next

	con.close: set con = nothing
	
	'
	' Al terminar, mostramos el resultado con la pagina de consulta!
	'
	
	Response.Redirect "ver.asp?t=" & Tiquete
%>