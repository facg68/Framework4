<%
    if Request.Cookies("usuario") = "" then
        Response.Redirect "/default.asp"
    else
        Response.Redirect "/forma/mobile"
    end if    
%>
