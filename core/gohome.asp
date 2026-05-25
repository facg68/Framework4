<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
			function HomePage(fxUsuario)
				dim t, cmdString, conn

                set conn = Server.CreateObject("ADODB.Connection")
                conn.open Application("Conn")
					cmdString = "SELECT usuCodigo, usuHomePage, ISNULL(LEN(usuHomePage), 0) AS Tam " & _
								"FROM seg_Usuarios " & _
								"WHERE usuCodigo = '" & fxUsuario & "';"

					set t = conn.execute(cmdString)
						if not (t.bof or t.eof) then
							if t("Tam") > 0 then
								HomePage = "/forma/desktop/apps/" & t("usuHomePage") & ".asp"
							else
								HomePage = "/core/desktop.asp"
							end if
						else
							HomePage = Application("DefPage")
						end if
					t.close: set t = nothing
                conn.close: set conn = nothing
			end function
        %>
    </head>

    <body>
        <%
            response.redirect HomePage(Request.Cookies("usuario"))
        %>
    </body>
</html>