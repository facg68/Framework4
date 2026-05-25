<%
    if Request.Cookies("usuario") = "" then
        Response.Redirect "/default.asp"
    else
        Response.Redirect "/core"
    end if    
%>
