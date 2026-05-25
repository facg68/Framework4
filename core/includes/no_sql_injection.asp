<%
	PaginaDeError = "/core/includes/sql_injection.asp?"

	no_SQL_injection_ListaNegra = Array("--", ";", "/*", "*/", "@@", "@", "char", "nchar", "varchar", "nvarchar",_
					"alter", "begin", "cast", "create", "cursor", "declare", "delete", "drop", "end", "exec",_
					"execute", "fetch", "insert", "kill", "open", "select", "sys", "sysobjects", "syscolumns",_
					"table", "update", "xp_", "<script")

	Function VerificarCadenaSQL(str) 
		'
		' Esta Funcion se carga automáticamente y es usada de
		' inmediato por el sistema
		'
		Dim lstr 	
		
		If ( IsEmpty(str) ) Then
			VerificarCadenaSQL = false
			Exit Function
		ElseIf ( StrComp(str, "") = 0 ) Then
			VerificarCadenaSQL = false
			Exit Function
		End If
	
		lstr = LCase(str)

		' Verifica si la cadena contiene algun patron de nuestra lista negra

		VerificarCadenaSQL = false
		
		For Each s in no_SQL_injection_ListaNegra
			If ( InStr (lstr, s) <> 0 ) Then
				VerificarCadenaSQL = true
			End If
		Next
	End Function 
	
	Function VerificarSQLInjection(Cadena)
		'
		' Devuelve un valor INT con la CANTIDAD de Posibles
		' Ataques del tipo SQL Injection que se encuentran
		' contenidos en la cadena en cuestion
		' 
		' Esta funcion esta de EJEMPLO por si un programador
		' decide no usar los servicios de segurida de InterWeb
		' pero no es necesaria si se usa el INCLUDE de esta
		' pagina en las que se desea proteger
		'
		
		dim con, sqlString, t, res
		
		sqlString = "SELECT dbo.Cripto_SQL_Injection('" & Cadena & "') AS R;"
		
		set con = server.CreateObject("ADODB.Connection")
		con.open Application("Conn")
		set t = con.execute(sqlString)
		
		VerificarSQLInjection = t("r")
		
		t.close: set t = nothing
		con.close: set con = nothing
	End Function 		

	'-----------------------------------------------
	'
	' Verificamos la Data enviada mediante FORMs
	'
	'-----------------------------------------------
	
	For Each s in Request.Form
		no_sql_injection_token = Request.Form(s)
		no_sql_injection_resultado = VerificarCadenaSQL(no_sql_injection_token)
		no_sql_injection_resultado_int = int(no_sql_injection_resultado)
		
		'response.write "FORM [" & no_sql_injection_token & "] --> " & no_sql_injection_resultado & " == " & no_sql_injection_resultado_int & "<br>"
	
		If  ( no_sql_injection_resultado_int  <> 0 ) Then
			PaginaDeError = PaginaDeError & "k=" & no_sql_injection_token & "&t=FORM"
			response.Redirect(PaginaDeError)
		End If
	Next
	
	'---------------------------------------------------
	'
	' Verificamos la Data enviada mediante QUERYSTRING
	'
	'---------------------------------------------------

	For Each s in Request.QueryString
		no_sql_injection_token = Request.QueryString(s)
		no_sql_injection_resultado = VerificarCadenaSQL(no_sql_injection_token)
		no_sql_injection_resultado_int = int(no_sql_injection_resultado)
		
		'response.write "QUERYSTRING [" & no_sql_injection_token & "] --> " & no_sql_injection_resultado & " == " & no_sql_injection_resultado_int & "<br>"
	
		If  ( no_sql_injection_resultado_int  <> 0 ) Then
			PaginaDeError = PaginaDeError & "k=" & no_sql_injection_token & "&t=QUERYSTRING"
			response.Redirect(PaginaDeError)
		End If
	Next
	
	'-----------------------------------------------
	'
	' Verificamos la Data enviada mediante COOKIES
	'
	'-----------------------------------------------

	For Each s in Request.Cookies
		no_sql_injection_token = Request.Cookies(s)
		no_sql_injection_resultado = VerificarCadenaSQL(no_sql_injection_token)
		no_sql_injection_resultado_int = int(no_sql_injection_resultado)
		
		'response.write "COOKIES [" & no_sql_injection_token & "] --> " & no_sql_injection_resultado & " == " & no_sql_injection_resultado_int & "<br>"
	
		If  ( no_sql_injection_resultado_int  <> 0 ) Then
			PaginaDeError = PaginaDeError & "k=" & no_sql_injection_token & "&t=COOKIE"	
			response.Redirect(PaginaDeError)
		End If
	Next
	
	'-----------------------------------------------
	'
	' Aqui agregamos cualquier otro problema que 
	' querramos verificar antes de ejecutar nuestra
	' pagina
	'
	'-----------------------------------------------
%>	