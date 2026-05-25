<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 

<html>
    <head>
        <%
            dim c, sqlString, usu, cont, Categ

            set c = Server.CreateObject("ADODB.Connection")
            c.open Application("Conn")        

            Function QueTipo(Categ)
                dim sqlCommand, tt

                sqlCommand = "SELECT Tipo FROM con_Contactos_Categorias " & _
                              "WHERE Usuario = '" & usu & "' " & _
                                "AND Codigo='" & Categ & "';"
                set tt = c.execute(sqlCommand)
                    if not (tt.bof or tt.eof) then
                        QueTipo = tt("Tipo")
                    end if
                tt.close: set tt = nothing
            End Function
        %>
    </head>

    <body>
        <%
            usu = Request.Cookies("Usuario")
            cont = Request.QueryString("c")
            categ = Request.QueryString("k")
            tipo = Request.QueryString("t")

            if tipo = "" then tipo = QueTipo(categ)
            
            sqlString = "DELETE FROM con_Contactos_ConCategs " & _
                         "WHERE (Usuario = '" & usu & "') " & _
                           "AND (Codigo = '" & cont & "') " & _                         
                           "AND (Tipo = '" & Tipo & "') " & _
                           "AND (Categoria = '" & categ & "');"

            c.execute sqlString
            c.close: set c = nothing

            response.redirect "cont_editar.asp?t=" & tipo & "&con=" & cont & "&tt=" & Request.QueryString("tt")
        %>    
    </body>
</html>