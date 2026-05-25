<%
    if Request.Cookies("usuario") = "" then
        Response.Redirect "login.asp"
    else
        Response.Redirect "/core"
    end if    
%>
