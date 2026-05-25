<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 

<%
    dim wp, cc, cmdString, NumWP

    NumWP = Request.QueryString("wp")
    Response.Cookies("usu_WP") = NumWP

    wp = RIGHT("0000000000" & NumWP, 8) & ".jpg"

    set cc = Server.CreateObject("ADODB.Connection")                      
    cc.open Application("Conn")  
        
        cmdString = "UPDATE seg_Usuarios SET usuWallPaper = '" & wp & "' WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';"
        cc.execute cmdString
    cc.close: set cc = nothing
%>