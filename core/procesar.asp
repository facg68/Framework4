<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    dim con, t, sqlBitacora, vinculo

    set con = Server.CreateObject("ADODB.Connection")
    con.open Application("Conn")
    set t = con.execute("SELECT proSistema, proAction FROM seg_Procesos WHERE proCodigo = '" & Request.QueryString("p") & "';")

    if not (t.eof or t.bof) then
        if Len(trim(t("proAction"))) > 0 then
            '
            ' Se ha definido una acción...
            ' Creamos el vínculo...
            '
            ' Aunque no se necesite, TODOS los vínculos llevan el sistema, el proceso y el usuario, aunque 
            ' se encuentre en el Cookie...
            '
            vinculo = "/forma/desktop/apps/" & t("ProSistema") & "/" & t("proAction") & ".asp?s=" & t("ProSistema") & "&p=" &  Request.QueryString("p") & "&u=" &  Request.Cookies("usuario")

            '
            ' Escribimos en la bitácora
            '
            sqlBitacora = "seg_pa_BitacoraAccesos '" & t("ProSistema") & "','" &  Request.QueryString("p") & "','" &  Request.Cookies("usuario") & "'"
            con.execute(sqlBitacora)


            '
            ' Abrimos la página...
            '
            response.Redirect vinculo
        else
            response.write "Este proceso aun no ha sido definido."
        end if
    end if

    t.close: set t = nothing
    con.close: set con = nothing
%>