<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function FechaServer(fecha)
                dim d, m, a, tiempo, ff

                if fecha <> "" then
                    ff = cStr(fecha)

                    d = left(ff, 2)
                    m = mid(ff, 4, 2)
                    a = mid(ff, 7, 4)
                    tiempo = right(ff, 5)

                    FechaServer = a & "-" & right("00" & m, 2) & "-" & right("00" & d, 2) & " " & tiempo
                else
                    FechaServer = NULL
                end if
            end function

            function Limpiar(valor)
                Limpiar = Replace(valor, "'", "´")
            end function                
        %>
    </head>

    <body>
        <%
            dim cc, tt, sqlString, usuario, secuencia, calendario, fsFecha, fsFechaFin
            dim Titulo, Fecha, FechaFin, TodoElDia, Repeticion, Direccion, Nota, Presupuesto, Monto, DbCr
            dim dia, mes, amo

            Usuario = Request.Cookies("Usuario")
            Origen =  Request.Form("Origen")
            fFiltro =  Request.Form("fFiltro")
            Dia =  Request.Form("Dia")
            Mes =  Request.Form("Mes")
            Amo =  Request.Form("Amo")

            Secuencia = Request.Form("Secuencia")
            Calendario = Request.Form("Calendario")
            Titulo = Limpiar(Request.Form("Titulo"))
            Fecha = Request.Form("Fecha")
            FechaFin = Request.Form("FechaFin")
            TodoElDia = Request.Form("TodoElDia")
            Repeticion = Request.Form("Repeticion")
            Direccion = Limpiar(Request.Form("Direccion"))
            Nota = Limpiar(Request.Form("Nota"))
            Presupuesto = Request.Form("Presupuesto")

            if Presupuesto = "1" then
                Monto = Request.Form("Monto")
                DbCr = Request.Form("DbCr")
            end if

            fsFecha = FechaServer(Fecha)
            fsFechaFin = FechaServer(FechaFin)


            if secuencia = "*" then 
                if Presupuesto = "1" then
                    sqlString = "INSERT INTO cal_Eventos(Usuario, Calendario, Fecha, FechaFin, Titulo, TodoElDia, Repeticion, Presupuesto, Monto, DbCr, Direccion, Nota) " & _
                                "VALUES('" & Usuario & "','" & Calendario & "', '" & fsFecha & "', '" & fsFechaFin & "', " & _ 
                                    "'" & Titulo & "', " & TodoElDia & ", " & Repeticion & ", " & Presupuesto & ", " & Monto & ", " & DbCr & ", " & _
                                    "'" & Direccion & "', '" & Nota & "');"
                else
                    sqlString = "INSERT INTO cal_Eventos(Usuario, Calendario, Fecha, FechaFin, Titulo, TodoElDia, Repeticion, Presupuesto, Direccion, Nota) " & _
                                "VALUES('" & Usuario & "','" & Calendario & "', '" & fsFecha & "', '" & fsFechaFin & "', " & _ 
                                    "'" & Titulo & "', " & TodoElDia & ", " & Repeticion & ", " & Presupuesto & ", '" & Direccion & "', '" & Nota & "');"
                end if

            else
                if Presupuesto = "1" then
                    sqlString = "UPDATE cal_Eventos " & _
                                "SET Calendario = '" & Calendario & "'," & _ 
                                    " Fecha = '" & fsFecha & "'," & _ 
                                    " FechaFin = '" & fsFechaFin & "'," & _ 
                                    " Titulo = '" & Titulo & "'," & _ 
                                    " TodoElDia = " & TodoElDia & "," & _ 
                                    " Repeticion = " & Repeticion & "," & _ 
                                    " Presupuesto = " & Presupuesto & "," & _ 
                                    " Monto = " & Monto & "," & _ 
                                    " DbCr = " & DbCr & "," & _ 
                                    " Direccion = '" & Direccion & "'," & _ 
                                    " Nota = '" & Nota & "' " & _ 
                                "WHERE (Secuencia = " & secuencia & ");"
                else
                    sqlString = "UPDATE cal_Eventos " & _
                                "SET Calendario = '" & Calendario & "'," & _ 
                                    " Fecha = '" & fsFecha & "'," & _ 
                                    " FechaFin = '" & fsFechaFin & "'," & _ 
                                    " Titulo = '" & Titulo & "'," & _ 
                                    " TodoElDia = " & TodoElDia & "," & _ 
                                    " Repeticion = " & Repeticion & "," & _ 
                                    " Presupuesto = " & Presupuesto & "," & _ 
                                    " Direccion = '" & Direccion & "'," & _ 
                                    " Nota = '" & Nota & "' " & _ 
                                "WHERE (Secuencia = " & secuencia & ");"                
                end if
            end if

            set cc = Server.CreateObject("ADODB.Connection")
            cc.open Application("Conn")
                cc.execute(sqlString)
            cc.close: set cc = nothing

            select Case origen
                case "d": response.redirect "cal_eventos.asp?d=" & dia & "&m=" & mes & "&a=" & amo
                case "s": response.redirect "cal_semanal.asp?f=" & fFiltro
                case "m": response.redirect "cal_calendario.asp?m=" & mes & "&a=" & amo
            end select


response.write "Secuencia = " & secuencia & "<br/>"
response.write "Presupuesto = " & PResupuesto & "<br/>"
response.write sqlString       

response.write "<br /><br />"

response.write "Origen =" &  Request.Form("Origen")
response.write "<br />Filtro = " & Request.Form("fFiltro")
response.write "<br />Dia = " & Request.Form("Dia")
response.write "<br />Mes = " & Request.Form("Mes")
response.write "<br />Amo = " & Request.Form("Amo")


        %>
    </body>
</html>