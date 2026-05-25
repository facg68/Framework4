<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>
<html>
    <head>
        <style>
            html, body {
                margin: 0;
                padding: 0;
                overflow: hidden;
                background-color: black;
                height: 100%;
            }     
            
            embed {
            width: 100%;
            height: 100%;
            border: none;
            }            
        </style>

        <%
            dim con, t, sqlString, anuncio
            Secuencia = request.querystring("p")
        %>
    </head>

    <%
        set con = Server.CreateObject("ADODB.Connection")
        con.open Application("Conn")

        sqlString = "SELECT Objeto FROM seg_Panfletos " & _
                    "WHERE (secuencia = " & Secuencia & ");"            

            set t = con.execute(sqlString)
                response.write "<body style='background-color: rgb(0, 0, 0);'>"
                    response.write "<embed src='/forma/desktop/apps/anuncios/pdf/" & t("Objeto") & "' type='application/pdf' />"
                response.write "</body>"
            response.write "</body>"   

        t.close: set t = nothing    
        con.close: set con = nothing    
    %>
</html>