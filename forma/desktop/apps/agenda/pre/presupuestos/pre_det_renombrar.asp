<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function Existe(Usuario, Codigo)
                dim con, tt, sqlString

                sqlstring = "select Count(*) AS Cuantos " & _
                            "from pre_Presupuesto_Encabezado " & _
                            "where (Usuario = '" & Usuario & "') " & _
                            "and (presupuesto = '" & Presupuesto & "');"

                set con = server.CreateObject("ADODB.Connection")
                con.open Application("Conn")
                set tt = con.execute(sqlString)

                if (tt.bof or tt.eof) then
                    Existe = 0
                else 
                    if tt("Cuantos") = 0 then   
                        Existe = 0
                    else
                        Existe = 1
                    end if
                end if

                tt.close: set tt = nothing
                con.close: set con = nothing
            end function
        %>
    </head>

    <body>
        <%
            dim c, pre, usu, sqlString, eNuevo
            dim t, v, e, o, nCod, nNom

            pre = Request.QueryString("p")
            usu = Request.Cookies("Usuario")

            d = Request.QueryString("d")
            t = Request.QueryString("t")
            v = Request.QueryString("v")
            e = Request.QueryString("e")
            o = Request.QueryString("o")

            nCod = Request.QueryString("cod")
            nNom = Request.QueryString("nom")

            if Existe(usu, nCod) = 0 then

                sqlString = "UPDATE pre_Presupuesto_Encabezado " & _
                                "SET Presupuesto = '" & nCod & "', " & _
                                   " Nombre = '" & nNom & "' " & _
                             " WHERE (Usuario = '" & usu & "') " & _
                                "AND (Presupuesto = '" & pre & "');"

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    c.execute (sqlString)
                c.close: set c = nothing

            end if

            response.redirect "pre_det_editar.asp?p=" & nCod & "&d=" & d & "&v=" & v & "&t=" & t & "&e=" & e & "&o=" & o
        %>
    </body>
</html>