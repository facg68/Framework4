        <div class="menu-toggle" onclick="toggleMenu()">
            <div class="bar"></div>
            <div class="bar"></div>
            <div class="bar"></div>
        </div>

        <nav class="sidebar" id="sidebar">
            <span class="close-btn" onclick="toggleMenu()">✕</span>

            <br />

            <%
                mob_sqlString = "exec dbo.seg_movil_menubar '" & Request.Cookies("Usuario") & "';"

                set mob_table = mob_conn.execute(mob_sqlString)
                    if not (mob_table.bof or mob_table.eof) then
                        Do
                            if (left(uCase(mob_table("proAction")), 4) = "HTTP") or (left(uCase(mob_table("proAction")), 3) = "WWW") then
                                vinculo = mob_table("proAction")
                            else
                                vinculo = "/forma/mobile/apps/" & mob_table("sistema") & "/"  &mob_table("proAction") & ".asp"
                            end if

                            response.write "<a href= '" & vinculo & "'"
                                if mob_table("Externo") = 1 then
                                    response.write " target='_blank' rel='noopener noreferrer'"
                                end if                          
                            response.write ">" & mob_table("proNombre") & "</a>"

                            mob_table.MoveNext
                        Loop Until (mob_table.eof)
                    end if
                mob_table.close: set mob_table = nothing
            %>
            
            <a href="/forma/mobile/login.asp">Salir</a>
        </nav>

        <header>
            <img src="/forma/mobile/recursos/images/logo_movil_blanco.png"  style="height: 35px; padding: 0;"/>
        </header>

        <div class="plus-overlay" onclick="cerrarObjetosPlus()"></div>
    <% mob_conn.close: set mob_conn = nothing %>