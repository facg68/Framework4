<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    tipo = Request.QueryString("tipo")

    set con = Server.CreateObject("ADODB.Connection")
    con.open Application("Conn")

    sql = "exec discos_filtro_Formas '" & Request.Cookies("Usuario") & "','" & tipo & "'"

    set rs = con.execute(sql)
        if not (rs.bof or rs.eof) then
            do while not rs.eof
                Response.Write "<option value='" & rs("Forma") & "'>"
                    Response.Write rs("Nombre")
                Response.Write "</option>"

                rs.movenext
            loop
        end if
    rs.close: con.close
%>