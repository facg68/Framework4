<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            function FechaHora2Fecha(fechaHora)
                dim dia, mes, amo

                amo = mid(fechaHora, 7, 4)
                mes = mid(fechahora, 4, 2)
                dia = left(fechaHora, 2)

                FechaHora2Fecha = amo & mes & dia
            end function

            function FechaHora2Hora(fechaHora)
                dim hh, mm

                hh = mid(fechaHora, 12,2)
                mm = mid(fechaHora, 15,2)

                FechaHora2Hora = hh & mm
            end function     

            Function SiVacio(valor)
                if (len(trim(valor)) = 0) OR (valor = NULL) then
                    SiVacio = "NULL"
                else
                    SiVacio = valor
                end if
            end Function       
        %>
    </head>

    <body>
        <%
            dim con, t, sqlString, fecha, hora, Penales
            dim etapa, grupo, fechahora, equipo1, goles1, equipo2, goles2, secuencia

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            secuencia = Request.Form("frmSec")
            etapa = Request.Form("frmEtapa")
            grupo = UCase(Request.Form("frmGrupo"))
            fechahora = Request.Form("frmFechaHora")
            equipo1 = Request.Form("frmEquipo1")
            goles1 = Request.Form("frmGoles1")
            equipo2 = Request.Form("frmEquipo2")
            goles2 = Request.Form("frmGoles2")
            Penales = Request.Form("frmPenales")

            fecha = FechaHora2Fecha(fechahora)
            hora = FechaHora2Hora(fechahora)

            if secuencia <> "" then
                sqlString = "UPDATE Mundial_Resultados " & _ 
                            "SET Etapa = '" & Etapa & "', " & _
                               " Grupo = '" & Grupo & "', " & _
                               " Fecha = '" & Fecha & "', " & _
                               " Hora = '" & Hora & "', " & _
                               " Equipo1 = '" & Equipo1 & "', " & _
                               " Goles1 = " & SiVacio(Goles1) & ", " & _
                               " Equipo2 = '" & Equipo2 & "', " & _
                               " Goles2 = " & SiVacio(Goles2) & ", " & _
                               " Penales = " & Penales & _
                         " WHERE Secuencia = " & Secuencia & ";"

                con.execute sqlString
            else
                sw = true

                if len(trim(etapa)) = 0 then sw = false
                if len(trim(grupo)) = 0 then sw = false
                if len(trim(fechahora)) = 0 then sw = false
                if len(trim(equipo1)) = 0 then sw = false
                if len(trim(equipo2)) = 0 then sw = false

                if sw then
                    sqlString = "INSERT INTO Mundial_Resultados(Etapa, Grupo, Fecha, Hora, Equipo1, Goles1, Equipo2, Goles2, Penales) " & _
                                "VALUES('" & Etapa & "', '" & Grupo & "', '" & Fecha & "', '" & Hora & "', '" & Equipo1 & "', " & SiVacio(Goles1) & ", '" & Equipo2 & "', " & SiVacio(Goles2) & ", " & Penales & ");"

                    con.execute sqlString
                end if
            end if

            con.close: set con = nothing

            response.redirect "historial.asp"
        %>
    </body>
</html>