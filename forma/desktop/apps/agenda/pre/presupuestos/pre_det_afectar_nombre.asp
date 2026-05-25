<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function Existe(Usuario, Presupuesto)
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
            dim c, sqlString, usuario, presupuesto, nombre, codigo 

            usuario = Request.Cookies("Usuario")
            presupuesto = Request.QueryString("p")
            nombre = Request.QueryString("n")

            if Existe(usuario, presupuesto) = 1 then
                sqlString = "UPDATE pre_Presupuesto_Encabezado " & _
                                "SET Nombre = '" & nombre & "'  " & _
                             " WHERE (Usuario = '" & usuario & "') " & _
                                "AND (Presupuesto = '" & presupuesto & "');"

                set c = server.CreateObject("ADODB.Connection")
                c.open Application("Conn")
                    c.execute (sqlString)
                c.close: set c = nothing
            end if
        %>
    </body>
</html>