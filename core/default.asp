<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<head>
    <meta charset="utf-8" />
    <%
        function HomePage()
            dim cc, t, cmdString

            cmdString = "SELECT usuHomePage FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';"

            set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")

                set t = cc.execute(cmdString)
                    if not (t.bof or t.eof) then
                        if len(trim(t("usuHomePage"))) > 0 then
                            HomePage = "/forma/desktop/apps/" & t("usuHomePage") & ".asp"
                        else
                            HomePage = "/core/desktop.asp"
                        end if
                    else
                        HomePage = Application("DefPage")
                    end if
                t.close: set t = nothing
            cc.close: set cc = nothing
        end function  
    %>
    </head>

    <body>
        <%
            if Request.Cookies("usuario") = "" then
                Response.Redirect "/default.asp"      
            else
                response.redirect HomePage()
            end if    
        %>
    </body>
</html>