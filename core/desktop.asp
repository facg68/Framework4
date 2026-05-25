<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">

        <!-- #include virtual = "/core/includes/kernel/head.inc" -->    

        <%
            dim Desktop_Conn, Indice, CuantosSnippets, Snippets(10, 2)

            set Desktop_Conn = Server.CreateObject("ADODB.Connection")                      
            Desktop_Conn.open Application("Conn")       

            Function ArchivoExiste(archivo)
                Dim fso
                Set fso = Server.CreateObject("Scripting.FileSystemObject")
                
                If fso.FileExists(Server.MapPath(archivo)) Then
                    ArchivoExiste = 1
                Else
                    ArchivoExiste = 0
                End If
                
                Set fso = Nothing
            End Function                   

            function wallPaper()
                dim t, cmdString

                cmdString = "SELECT usuWallPaper FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';"
                wallpaper = "00000001.jpg"

                set t = Desktop_Conn.execute(cmdString)

                if not (t.bof or t.eof) then
                    if len(trim(t("usuWallPaper"))) > 0 then
                        wallpaper = "/core/imagenes/fondos/" & t("usuWallPaper")
                    end if
                end if

                t.close: set t = nothing
            end function     

            function random_wallPaper()
                dim wallPaper

                wallPaper = RandomNumber(Request.Cookies("max_WP"))
                wallPaper = right("00000000" & wallPaper, 8) & ".jpg"
                random_wallPaper = "/core/imagenes/fondos/" & wallPaper
            end function                

            function SnippetsActivados()
                dim t, cmdString

                cmdString = "SELECT usuCargarSnippets FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';"

                set t = Desktop_Conn.execute(cmdString)

                if not (t.bof or t.eof) then
                    SnippetsActivados = t("usuCargarSnippets")
                else
                    SnippetsActivados = 0
                end if

                t.close: set t = nothing
            end function   

            function IniciarMinimizado()
                dim t, cmdString

                cmdString = "SELECT usuIniciarSinEncabezado FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';"

                set t = Desktop_Conn.execute(cmdString)

                if not (t.bof or t.eof) then
                    IniciarMinimizado = t("usuIniciarSinEncabezado")
                else
                    IniciarMinimizado = 0
                end if

                t.close: set t = nothing            
            end function     

            function desktop_wallPaper()
                dim t, cmdString

                cmdString = "SELECT usuRandomWallpaper FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';"

                set t = Desktop_Conn.execute(cmdString)

                if not (t.bof or t.eof) then
                    if t("usuRandomWallpaper") = 1 then
                        desktop_wallPaper = random_wallPaper()
                    else
                        desktop_wallPaper = WallPaper()
                    end if                    
                else
                    desktop_wallPaper = WallPaper()
                end if

                t.close: set t = nothing            
            end function 

            function desktop_vWallPaper()
                 desktop_vWallPaper = desktop_WallPaper()
                 desktop_vWallPaper =  replace(desktop_vWallPaper, "jpg", "mp4")
            end function
        %>

        <style>
            html, body {
                height: 100%;
                margin: 0;
                padding: 0;
            }
                    
            .desktop_container_menu {
                height:100vh;
                position: absolute;
                top: 65px;
                right: 0px;
                left: 0px;
                bottom: 0px;
                z-index: 1;
            }            

            .desktop_container_no_menu {
                height:100vh;
                position: absolute;
                top: 21px;
                right: 0px;
                left: 0px;
                bottom: 0px;
                z-index: 2;                
            }     
            
            .bg {
                background-image: url("<%= desktop_wallPaper() %>");

                height: 100%;
                background-position: center;
                background-repeat: no-repeat;
                background-size: cover;
            }

            #myVideo {
                position: relative;
                right: 0;
                bottom: 0;
                min-width: 100%; 
                min-height: 100%;
                z-index: 0;                
            }

            .content {
                position: fixed;
                bottom: 0;
                background: rgba(0, 0, 0, 0.5);
                color: #f1f1f1;
                width: 100%;
                padding: 20px;
            }   
            
            #desk_taskbar {
                position: fixed;
                bottom: 0;
                left: 0;
                right: 0;

                height: 34px;

                transition: opacity 0.3s ease;
                opacity: 0;
                display: none;

                align-items: center;
                gap: 4px;
                padding: 4px;

                background: linear-gradient(
                    to top,
                    rgba(40,40,40,0.45),
                    rgba(80,80,80,0.15)
                );

                backdrop-filter: blur(6px);
                -webkit-backdrop-filter: blur(6px);
                z-index: 9999;
            }
            
            .taskbar-tile {
                background: hsla(0, 0%, 28%, 0.13);
                border: 1px solid rgba(255,255,255,0.06);
                border-radius: 6px;

                color: #ffffff;
                padding: 3px 12px;
                font: 12px sans-serif;
                transition: all 0.15s ease;

                backdrop-filter: blur(4px);
                -webkit-backdrop-filter: blur(4px);

                cursor: pointer;
            }

            .taskbar-tile:hover {
                background: rgba(206, 206, 206, 0.85);
                color: black;
                transform: translateY(-1px);
                box-shadow: 0 0 6px rgba(255,255,255,0.3);
            }            
        </style>        
    </head>

    <body class="bg" plantilla="desktop" reserva="0" 
          plus="mbSnippets: plus_mbSnippets @ mouseenter | lista ;"
          onload="init(<%= SnippetsActivados() %>)">

        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <div id="desk_div_container" style="overflow: hidden;" 
             class="<%
                if IniciarMinimizado() = 1 then
                    response.write "desktop_container_no_menu"
                else
                    response.write "desktop_container_menu"
                end if        
            %>">

            <% if ArchivoExiste(desktop_vWallPaper()) = 1 then %>
                <video autoplay muted loop id="myVideo">
                    <source src="<%= desktop_vWallPaper() %>" type="video/mp4">
                </video>        
            <% end if %>           
        </div> 

        <div id="desk_taskbar">
        </div>        

        <!-- #include virtual = "/core/includes/kernel/winman.inc" -->

        <script>
            mb_Maximizar = () => Desktop_Maximizar();
            mb_Minimizar = () => Desktop_Minimizar();            

            function Desktop_Minimizar() {
                var miDiv = document.getElementById("desk_div_container");
                miDiv.setAttribute("class", "desktop_container_no_menu");
            }

            function Desktop_Maximizar() {
                var miDiv = document.getElementById("desk_div_container");
                miDiv.setAttribute("class", "desktop_container_menu");
            }        
        </script>
    </body>

    <% Desktop_Conn.close: set Desktop_Conn = nothing %>

    <!-- #include virtual = "/forma/plus/plusSnippets.plus" -->
    <!-- #include virtual = "/core/includes/kernel/close.inc" -->
</html>