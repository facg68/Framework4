        <meta http-equiv="Content-Type"     content="text/html; charset=UTF-8"/>
        <meta http-equiv="X-UA-Compatible"  content="ie=edge" />  

        <meta name="viewport"               content="width=device-width, initial-scale=1.0">
        <meta name="extranet"               content="">
        <meta name="fabrizio"               content="Dashboard">
        <meta name="keyword"                content="apps, calendar, contacts, media library">
        
        <title><%= Application("Nombre") %></title>
        
        <link href="/core/imagenes/apple-touch-icon.png"  rel="icon">
        <link href="/core/imagenes/apple-touch-icon.png"  rel="apple-touch-icon">

        <link href="/core/lib/font-awesome-6/all.min.css" rel="stylesheet" />
        <link href="/core/lib/swal-2/sweetalert2.min.css" rel="stylesheet">     

        <script src="/core/lib/swal-2/sweetalert2.all.min.js"></script>   
        <script src="/core/lib/apex/apexcharts.min.js"></script>                    

        <%
            dim mb_conn, mb_ListaSnippets(6, 3), thisSystem, thisProcess, thisNoMenu

            if Request.Cookies("usuario") = "" then
                Response.Redirect "/forma/desktop/login.asp"      
            end if

            set mb_conn = Server.CreateObject("ADODB.Connection")                      
            mb_conn.open Application("Conn")     

            ' Auxiliares ------------------------------------------------------------

                Sub SysLockOut()
                    dim seg_conn, t, cmdString, sis, proc

                    sis = thisSystem
                    proc = thisProcess

                    if isnull(Request.Cookies("Usuario")) or (Request.Cookies("Usuario") = "") then 
                        response.redirect "/forma/desktop/login.asp"
                    else
                        if isnull(sis) or (sis = "") then sis = "/*"
                        if isnull(proc) or (proc = "") then proc = "/*"          

                        cmdString = "SELECT COUNT(*) AS Acceso " & _
                                    "FROM seg_PermisosUsuarios " & _
                                    "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                    "AND (Sistema = '" & sis & "') " & _
                                    "AND (Proceso = '" & proc & "');"

                        set seg_conn = Server.CreateObject("ADODB.Connection")
                        seg_conn.open Application("Conn") 

                        set t = seg_conn.execute(cmdString)
                            if t("Acceso") = 0 then
                                response.redirect "/forma/desktop/login.asp"
                            end if            
                        t.close: set t = nothing

                        seg_conn.close: set seg_conn = nothing
                    end if
                end Sub	

                function HomePage()
                    dim t, cmdString

                    cmdString = "SELECT usuHomePage FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';"

                    set t = mb_conn.execute(cmdString)

                    if not (t.bof or t.eof) then
                        if len(trim(t("usuHomePage"))) > 0 then
                            HomePage = "/forma/desktop/" & t("usuHomePage") & ".asp"
                        else
                            HomePage = "/core/desktop.asp"
                        end if
                    else
                        HomePage = Application("DefPage")
                    end if

                    t.close: set t = nothing
                end function

                Function BitacoraActivada()
                    dim seg_conn, t, cmdString

                    cmdString = "SELECT sysBitacora " & _
                                    "FROM seg_Sistemas " & _
                                "WHERE (sysCodigo = '" & thisSystem & "');"

                    set seg_conn = Server.CreateObject("ADODB.Connection")
                    seg_conn.open Application("Conn") 
                        set t = seg_conn.execute(cmdString)
                            BitacoraActivada = t("sysBitacora")
                        t.close: set t = nothing    
                    seg_conn.close: set seg_conn = nothing                      
                end function
                    
                sub Bitacora()
                    dim seg_conn, t, sqlBitacora        

                    sqlBitacora = "seg_pa_BitacoraAccesos '" & thisSystem & "','" & thisProcess & "','" &  Request.Cookies("usuario") & "'"

                    if BitacoraActivada() = 1 then
                        set seg_conn = Server.CreateObject("ADODB.Connection")
                        seg_conn.open Application("Conn") 
                        seg_conn.execute sqlBitacora
                        seg_conn.close: set seg_conn = nothing     
                    end if
                end sub

                Function RandomNumber(Limite)
                    Randomize
                    RandomNumber = Int((Limite * Rnd) + 1)
                End Function   

                function unFormatNumber(Valor)
                    dim v 

                    v = trim(Valor)

                    v = Replace(v, Chr(160), "")        ' NBSP
                    v = Replace(v, ChrW(8239), "")      ' Thin space
                    v = Replace(v, ",", "")             ' miles
                    If Not IsNumeric(v) Then v = 0 

                    unFormatNumber = v           
                end function   

                Function formUnpack(campo, paquete)
                    Dim buscar, inicio, fin, valor
                    
                    buscar = "[" & ucase(campo) & "="
                    inicio = InStr(ucase(paquete), buscar)
                    
                    If inicio > 0 Then
                        inicio = inicio + Len(buscar)
                        fin = InStr(inicio, paquete, "]")
                        
                        If fin > 0 Then
                            valor = Mid(paquete, inicio, fin - inicio)
                        Else
                            valor = ""
                        End If                
                    Else
                        valor = ""
                    End If
                    
                    formUnpack = valor
                End Function                                             

            ' Fin - Auxiliares ------------------------------------------------------
      
            ' Menu Bar y Desktop ----------------------------------------------------

                Function NoVerEncabezado()
                    dim t, cmdString

                    NoVerEncabezado = 0
                    cmdString = "SELECT usuIniciarSinEncabezado FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';"

                    set t = mb_conn.execute(cmdString)

                    if not (t.bof or t.eof) then
                        if t("usuIniciarSinEncabezado") = 1 then
                            NoVerEncabezado = 1
                        end if
                    end if

                    t.close: set t = nothing    
                end function

                sub DibujarMenuBar(Usuario)
                    '
                    ' Dibujamos los Items "Root" que el usuario
                    ' puede ver (Menú Raíz)
                    '
                    dim t, cadena, cmdString

                    cmdString  = "SELECT * FROM seg_WebMenu WHERE (Usuario = '" & Usuario & "') ORDER BY Indice;"
                    set t = mb_conn.execute(cmdString)

                    if not (t.bof or t.eof) then
                        do
                            cadena = Replace(t("Comando"), "|", CHR(39))
                            response.write cadena

                            t.movenext
                        loop until t.eof
                    end if

                    t.close: set t = nothing
                end sub

                function EsAnclable()
                    dim seg_conn, t, cmdString, sis, proc

                    sis = thisSystem
                    proc = thisProcess

                    if isnull(sis) or (sis = "") then sis = "/*"
                    if isnull(proc) or (proc = "") then proc = "/*"   

                    cmdString = "SELECT ISNULL(Shortcut, 0) AS Estatus " & _
                                    "FROM seg_Procesos " & _
                                "WHERE (proSistema = '" & sis & "') " & _
                                    "AND (proCodigo = '" & proc & "');"

                    set seg_conn = Server.CreateObject("ADODB.Connection")
                    seg_conn.open Application("Conn") 
                        set t = seg_conn.execute(cmdString)
                            if not (t.bof or t.eof) then
                                EsAnclable = t("Estatus")
                            else
                                EsAnclable = 0
                            end if   
                        t.close: set t = nothing
                    seg_conn.close: set seg_conn = nothing           
                end function

                function EstaAnclado()
                    dim seg_conn, t, cmdString

                    if isnull(thisSystem) or (thisSystem = "") then thisSystem = "/*"
                    if isnull(thisProcess) or (thisProcess = "") then thisProcess = "/*"

                    cmdString = "SELECT COUNT(*) AS Anclado " & _
                                    "FROM seg_Usuarios_Shortcuts " & _
                                "WHERE (codUsuario = '" & Request.Cookies("Usuario") & "') " & _
                                    "AND (codSistema = '" & thisSystem & "') " & _
                                    "AND (codProceso = '" & thisProcess & "');"

                    set seg_conn = Server.CreateObject("ADODB.Connection")
                    seg_conn.open Application("Conn") 
                        set t = seg_conn.execute(cmdString)
                            if not (t.bof or t.eof) then
                                EstaAnclado = t("Anclado")
                            else
                                EstaAnclado = 0
                            end if   
                        t.close: set t = nothing
                    seg_conn.close: set seg_conn = nothing        
                end function        

                Sub PresentarShortcuts()
                    dim seg_conn, t, cmdString

                    cmdString = "SELECT p.proNombre, p.proIcon, p.proAction, '/forma/desktop/apps/' + s.codSistema + '/' + p.proAction + '.asp' AS Accion " & _
                                    "FROM seg_Usuarios_Shortcuts AS s " & _
                            "INNER JOIN seg_Procesos AS p " & _
                                    "ON s.codSistema = p.proSistema " & _
                                    "AND s.codProceso = p.proCodigo " & _
                                "WHERE (s.codUsuario = '" & Request.Cookies("Usuario") & "') " & _
                                "ORDER BY p.proNombre"

                    set seg_conn = Server.CreateObject("ADODB.Connection")
                    seg_conn.open Application("Conn") 
                        set t = seg_conn.execute(cmdString)
                            if not (t.bof or t.eof) then
                                do
                                %>
                                    <img class="mb_icon" style="filter: none;"  
                                        src="/core/imagenes/<%= t("proIcon") %>" 
                                        height="35" width="35" 
                                        onclick="abrirShortcut('<%= t("Accion") %>')" alt="<%= t("proNombre") %>">           
                                <%
                                t.movenext
                                Loop until t.eof
                            end if
                        t.close: set t = nothing
                    seg_conn.close: set seg_conn = nothing    
                end sub                

                Sub IconosMB()
                    %>
                        <img class="mb_icon" style="filter: none;" id="mbSnippets"   
                             src="/core/imagenes/desktop.png" 
                             height="35" width="35" 
                             onclick="escritorio()" alt="Modo Escritorio">

                        <img class="mb_icon" style="filter: none;" id="perfil_icon"   
                             src="/core/imagenes/perfil.png" 
                             height="35" width="35" >
                             
                        <img class="mb_icon" style="filter: none;"  
                             src="/core/imagenes/buscar.png" 
                             height="35" width="35" 
                             onclick="buscar()" alt="Buscar Informacion">
                    <%

                    PresentarShortcuts

                    %>
                        <img class="mb_icon" style="filter: none;"  
                             src="/core/imagenes/notificaciones.png" 
                             height="35" width="35" 
                             onclick="abrirNotificaciones()" alt="Ver Lista de Notificaciones">
                             
                        <img class="mb_icon" style="filter: none;"  
                             src="/core/imagenes/salir.png" 
                             height="35" width="35" 
                             onclick="salir()" alt="Cerrar la Aplicación y Salir del Sistema">
                    <%
                end Sub   

                function wallPaper()
                    dim t, cmdString

                    cmdString = "SELECT usuWallPaper FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';"
                    wallpaper = "00000001.jpg"

                    set t = cc.execute(cmdString)

                    if not (t.bof or t.eof) then
                        if len(trim(t("usuWallPaper"))) > 0 then
                            wallpaper = t("usuWallPaper")
                        end if
                    end if

                    t.close: set t = nothing
                end function       

                Sub ActualizarWallPaper(indice)
                    dim cmdString, wpCadena

                    wpCadena = right("0000000000" & indice, 8) & ".jpg"
                    cmdString = "UPDATE seg_Usuarios SET usuWallPaper = '" & wpCadena & "' WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';"

                    cc.execute cmdString
                end Sub                               

            ' Fin - Menu Bar ---------------------------------------------------------    

            ' Snippets ---------------------------------------------------------------

                function NombrePagina()
                    dim files, url, segmentos

                    url = Request.ServerVariables("path_info")
                    segmentos = split(url,"/")

                    'leer el ultimo segmento de la cadena URL
                    
                    url = segmentos(ubound(segmentos))
                    NombrePagina = url
                end function    

                function Opacidad()
                    dim t, cmdString

                    cmdString = "SELECT snippetsOpacidad FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';"

                    set t = mb_conn.execute(cmdString)

                    if not (t.bof or t.eof) then
                        Opacidad = ( t("snippetsOpacidad") / 100 ) 
                    else
                        Opacidad = 0.95
                    end if

                    t.close: set t = nothing
                end function                                          

            ' Fin - Snippets ---------------------------------------------------------

            ' Parametros de Usuario y Sistemas ---------------------------------------

                Function DefaultParametro(Sistema, Parametro)
                    '
                    ' Hay que volver a escribir este proceso
                    ' Hay que reflejar los 5 tipos de Parámetros
                    '
                    dim seg_conn, t, cmdString, sis, proc

                    cmdString = "SELECT ValorDefault  " & _
                                "FROM seg_Parametros " & _
                                "WHERE (Sistema = '" & Sistema & "') " & _
                                "AND (Parametro = '" & Parametro & "');"

                    set seg_conn = Server.CreateObject("ADODB.Connection")
                    seg_conn.open Application("Conn") 
                        set t = seg_conn.execute(cmdString)
                            if not (t.bof or t.eof) then
                                DefaultParametro = t("ValorDefault")
                            else
                                DefaultParametro = NULL
                            end if   
                        t.close: set t = nothing
                    seg_conn.close: set seg_conn = nothing               
                end Function

                Function ParametroUsuario(Sistema, Parametro)
                    dim seg_conn, t, cmdString, sis, proc

                    cmdString = "SELECT Valor " & _
                                "FROM seg_Usuarios_Parametros " & _
                                "WHERE (Sistema = '" & Sistema & "') " & _
                                "AND (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                "AND (Parametro = '" & Parametro & "');"

                    set seg_conn = Server.CreateObject("ADODB.Connection")
                    seg_conn.open Application("Conn") 
                        set t = seg_conn.execute(cmdString)
                            if not (t.bof or t.eof) then
                                ParametroUsuario = t("Valor")
                            else
                                ParametroUsuario = NULL
                            end if   
                        t.close: set t = nothing
                    seg_conn.close: set seg_conn = nothing               
                end Function

                Function PuedeAccesar(Sistema, Proceso)
                    dim seg_conn, t, cmdString, sis, proc

                    cmdString = "SELECT COUNT(*) AS Cuantos " & _
                                "FROM seg_PermisosUsuarios " & _
                                "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                "AND (Sistema = '" & Sistema & "') " & _
                                "AND (Proceso = '" & Proceso & "')"

                    set seg_conn = Server.CreateObject("ADODB.Connection")
                    seg_conn.open Application("Conn") 
                        set t = seg_conn.execute(cmdString)
                            if not (t.bof or t.eof) then
                                PuedeAccesar = t("Cuantos")
                            end if   
                        t.close: set t = nothing
                    seg_conn.close: set seg_conn = nothing               
                end Function

            ' Fin - Parametros de Usuario y Sistemas ---------------------------------

            ' Framework Params -------------------------------------------------------

                Function framework_param(param, default)
                    framework_param = ParametroUsuario("seguridad", param)    

                    if isnull(framework_param) or (framework_param = "") then
                        framework_param =  DefaultParametro("seguridad", param)

                        if isnull(framework_param) or (framework_param = "") then
                            framework_param = default
                        end if
                    end if                  
                end Function

                '---------- Barra de Menues Principal ------------------------------------------

                    Function framework_mb_logo_color()
                        framework_mb_icon_color = framework_param("title_area_logo_color", "#ffffff")
                    end Function  

                    Function framework_mb_background()
                        framework_mb_background = framework_param("title_area_mb_color", "#000000")
                    end Function    

                    Function framework_mb_font()
                        framework_mb_font = ParametroUsuario("seguridad", "title_area_font")

                        if isnull(framework_mb_font) or (framework_mb_font = "") then 
                            framework_mb_font = DefaultParametro("seguridad", "title_area_font")

                            if isnull(framework_mb_font) or (framework_mb_font = "") then 
                                framework_mb_font = "Arial"
                            end if
                        end if            

                        framework_mb_font = Replace(framework_mb_font, "|", "'") 
                    end Function                                      

                    Function framework_mb_font_color()
                        framework_mb_font_color = framework_param("title_area_font_color", "#ffffff")
                    end Function  

                    Function framework_menu_font_size()
                        framework_menu_font_size = framework_param("title_area_font_size", "14px")
                    end Function    

                    Function framework_mb_icon_color()
                        framework_mb_icon_color = framework_param("title_area_icons_color", "#ffffff")
                    end Function  

                    Function framework_menu_drop_shadow()
                        framework_menu_drop_shadow = (framework_param("title_area_shadow", "30") / 100)
                    end Function                       

                '-------------------------------------------------------------------------------

                '---------- Paneles de Menú ----------------------------------------------------
 
                    Function framework_menu_background_color()
                        framework_menu_background_color = framework_param("menu_panel_color", "#ffffff")
                    end Function               

                    Function framework_menu_font_color()
                        framework_menu_font_color = framework_param("menu_panel_font_color", "#FFFFFF")
                    end Function    

                '-------------------------------------------------------------------------------

            ' Fin - Framework Params -------------------------------------------------                

            ' Apex -------------------------------------------------------------------

                sub apexColumns(Titulo, NombreDivChart, SQL_Query, Codigo, Valor, color, gfxReserva)
                    dim gLabel, gData, sw, gfxTable, gfxConn

                    set gfxConn = Server.CreateObject("ADODB.Connection")
                        gfxConn.open Application("Conn")
                        
                        set gfxTable = gfxConn.execute(SQL_Query)
                            if not (gfxTable.bof or gfxTable.eof) then
                                gLabel = ""
                                gData1 = ""
                                sw = 0

                                do
                                    if sw = 0 then
                                        sw = 1
                                        gLabel = gLabel & "['" & gfxTable(Codigo) & "'"
                                        gData1 = gData1 & "[" & gfxTable(Valor)
                                    else
                                        gLabel = gLabel & ", '" & gfxTable(Codigo) & "'"
                                        gData1 = gData1 & ", " & gfxTable(Valor)                                    
                                    end if

                                    gfxTable.MoveNext
                                loop until gfxTable.eof

                                gLabel = gLabel & "]"
                                gData1 = gData1 & "]"
                            end if
                        gfxTable.close: set gfxTable = nothing
                    gfxConn.close: set gfxConn = nothing

                    response.write "<div id='" & NombreDivChart & "' style='padding: 0px;'></div>"

                    %>
                        <script>
                            var options = {
                                series: [{
                                    name: '',
                                    data: <%= gData1 %>
                                }],

                                chart: {
                                    height: <%= gfxReserva %>,
                                    type: 'bar',
                                    toolbar: false,
                                },

                                title: {
                                    text: '<%= Titulo %>',
                                },                               

                                plotOptions: {
                                    bar: {
                                            borderRadius: 10,
                                            dataLabels: {
                                                position: 'top',
                                            },
                                    }
                                },
                                colors: ['<%= color %>'],
                                dataLabels: {
                                    enabled: true,                                     
                                    offsetY: -30,
                                    style: {
                                        fontSize: '12px',
                                        colors: ["<%= color %>"]
                                    }
                                },
                                
                                xaxis: {
                                    categories: <%= gLabel %>,
                                    position: 'bottom',
                                    axisBorder: {
                                            show: false
                                    },
                                    axisTicks: {
                                            show: false
                                    },
                                    crosshairs: {
                                        fill: {
                                            type: 'gradient',
                                            gradient: {
                                                colorFrom: '#D8E3F0',
                                                colorTo: '#BED1E6',
                                                stops: [0, 100],
                                                opacityFrom: 0.4,
                                                opacityTo: 0.5,
                                            }
                                        }
                                    },
                                
                                    tooltip: {
                                        enabled: true,
                                    }
                                },

                                yaxis: {
                                    axisBorder: {
                                        show: false
                                    },
                                    axisTicks: {
                                        show: false,
                                    },
                                    labels: {
                                        show: false,
                                    }        
                                }
                            };

                            var chart = new ApexCharts(document.querySelector("#<%= NombreDivChart %>"), options);
                            chart.render();
                        </script> 
                    <%           
                end sub  

                sub apexDonut(Titulo, NombreDivChart, SQL_Query, Codigo, Valor, gfxReserva)
                    dim gLabel, gData, sw, gfxTable, gfxConn

                    set gfxConn = Server.CreateObject("ADODB.Connection")
                        gfxConn.open Application("Conn")
                        
                        set gfxTable = gfxConn.execute(SQL_Query)
                            if not (gfxTable.bof or gfxTable.eof) then
                                gLabel = ""
                                gData1 = ""
                                sw = 0

                                do
                                    if sw = 0 then
                                        sw = 1
                                        gLabel = gLabel & "['" & gfxTable(Codigo) & "'"
                                        gData1 = gData1 & "[" & gfxTable(Valor)
                                    else
                                        gLabel = gLabel & ", '" & gfxTable(Codigo) & "'"
                                        gData1 = gData1 & ", " & gfxTable(Valor)                                    
                                    end if

                                    gfxTable.MoveNext
                                loop until gfxTable.eof

                                gLabel = gLabel & "]"
                                gData1 = gData1 & "]"
                            end if
                        gfxTable.close: set gfxTable = nothing
                    gfxConn.close: set gfxConn = nothing

                    response.write "<div id='" & NombreDivChart & "' style='padding: 0px;'></div>"
                    %>
                        <script>
                            var options = {
                                series: <%= gData1 %>,
                                chart: {
                                    height: <%= gfxReserva %>,
                                    type: 'donut',
                                },

                                labels: <%= gLabel %>,
                                plotOptions: {
                                    pie: {
                                        startAngle: -90,
                                        endAngle: 270
                                    }
                                },

                                dataLabels: {
                                    enabled: true,
                                    formatter: function (val) {
                                        return dataformat(val) + "%"
                                    },
                                },
                                
                                fill: {
                                    type: 'gradient',
                                },

                                title: {
                                    text: '<%= Titulo %>'
                                },

                                responsive: [{
                                    breakpoint: 480,
                                    options: {
                                        chart: {
                                            width: 200
                                        },
                                        legend: {
                                            position: 'bottom'
                                        }
                                    }
                                }]
                            };

                            var chart = new ApexCharts(document.querySelector("#<%= NombreDivChart %>"), options);
                            chart.render();                   
                        
                        </script>            
                    <%
                end sub  

                sub apexPanoramic(Titulo, SQL_Query, Codigo, Valor, Color, gfxReserva)
                    dim gLabel, gData, sw, gfxTable, gfxConn, vMax
                    vMax = 0.00

                    set gfxConn = Server.CreateObject("ADODB.Connection")
                        gfxConn.open Application("Conn")
                        
                        set gfxTable = gfxConn.execute(SQL_Query)
                            if not (gfxTable.bof or gfxTable.eof) then
                                gLabel = ""
                                gData1 = ""
                                sw = 0

                                do
                                    if sw = 0 then
                                        sw = 1
                                        gLabel = gLabel & "[" & gfxTable(Codigo) 
                                        gData1 = gData1 & "[" & gfxTable(Valor)
                                    else
                                        gLabel = gLabel & ", " & gfxTable(Codigo) 
                                        gData1 = gData1 & ", " & gfxTable(Valor)                                    
                                    end if

                                    if cDbl(gfxTable(Valor)) > vMax then vMax = cDbl(gfxTable(Valor))

                                    gfxTable.MoveNext
                                loop until gfxTable.eof

                                gLabel = gLabel & "]"
                                gData1 = gData1 & "]"
                            end if
                        gfxTable.close: set gfxTable = nothing
                    gfxConn.close: set gfxConn = nothing      
                 %>

                    <div>
                        <div id="chart-line2"></div>
                        <div id="chart-line"></div>
                    </div>

                    <script>
                        var chartHeight = <%= gfxReserva %>;
                        var BigChart = (chartHeight * 0.75);
                        var PanChart = (chartHeight - BigChart);

                        var serie = <%= gData1 %>;
                        var labels = <%= gLabel %>;

                        var options = {
                            series: [{
                                name: 'Valor',
                                data: <%= gData1 %>
                            }],
                            chart: {
                                id: 'bigChart',
                                type: 'area',
                                height: BigChart,               
                                toolbar: { show: false }
                            },
                            colors: ['<%= Color %>'],
                            stroke: { width: 3, curve: 'smooth' },
                            dataLabels: { enabled: false },
                            xaxis: { categories: <%= gLabel %> },
                            yaxis: { 
                                        title: { text: 'Valor' },
                                        min: 0,
                                        max: <%= vMax %>
                                },
                            fill: {
                                type: 'solid',
                                opacity: 1
                            },
                            title: {
                                text: '<%= Titulo %>',
                                align: 'left', // opciones: left, center, right
                                style: {
                                    fontSize: '20px',
                                    fontWeight: 'bold',
                                    color: '#111111ff'
                                }
                            }
                        };

                        var chart = new ApexCharts(document.querySelector("#chart-line2"), options);

                        var optionsLine = {
                            series: [{ name: 'Valor', data: <%= gData1 %> }],
                            chart: {
                                id: 'panoramicChart',
                                height: PanChart,
                                type: 'area',
                                brush: { target: 'bigChart', enabled: true },
                                selection: { enabled: true }
                            },
                            colors: ['<%= Color %>'],
                            stroke: { width: 1, curve: 'smooth' },
                            fill: {
                                type: 'gradient',
                                gradient: { opacityFrom: 0.91, opacityTo: 0.1 }
                            },
                            xaxis: {
                                categories: <%= gLabel %>,
                                tooltip: { enabled: false }
                            },
                            yaxis: {    
                                    max: <%= vMax %>, 
                                    min: 0,
                                    tickAmount: 4 
                                }
                        };

                        var chartLine = new ApexCharts(document.querySelector("#chart-line"), optionsLine);

                        Promise.all([chart.render(), chartLine.render()]).then(() => {
                            let total = labels.length;
                            let inicio = Math.max(0, total - 5) + 1;    // si hay menos de 20, arranca en 0
                            let fin = total;                            // último índice
                            chart.zoomX(inicio, fin);
                        });
                    </script>
                 <%          
                end sub      

                sub apexArea(Titulo, SQL_Query, Codigo, Valor, Color, gfxReserva)
                    dim gLabel, gData, sw, gfxTable, gfxConn, vMax
                    vMax = 0.00

                    set gfxConn = Server.CreateObject("ADODB.Connection")
                        gfxConn.open Application("Conn")
                        
                        set gfxTable = gfxConn.execute(SQL_Query)
                            if not (gfxTable.bof or gfxTable.eof) then
                                gLabel = ""
                                gData1 = ""
                                sw = 0

                                do
                                    if sw = 0 then
                                        sw = 1
                                        gLabel = gLabel & "[" & gfxTable(Codigo) 
                                        gData1 = gData1 & "[" & gfxTable(Valor)
                                    else
                                        gLabel = gLabel & ", " & gfxTable(Codigo) 
                                        gData1 = gData1 & ", " & gfxTable(Valor)                                    
                                    end if

                                    if cDbl(gfxTable(Valor)) > vMax then vMax = cDbl(gfxTable(Valor))

                                    gfxTable.MoveNext
                                loop until gfxTable.eof

                                gLabel = gLabel & "]"
                                gData1 = gData1 & "]"
                            end if
                        gfxTable.close: set gfxTable = nothing
                    gfxConn.close: set gfxConn = nothing      
                 %>

                    <div id="chart-line2"></div>

                    <script>
                        var chartHeight = (window.innerHeight - menuReserva - wBase - <%= gfxReserva %>);  
                        var serie = <%= gData1 %>;
                        var labels = <%= gLabel %>;

                        var options = {
                            series: [{
                                name: 'Valor',
                                data: serie
                            }],
                            chart: {
                                id: 'bigChart',
                                type: 'area',
                                height: <%= gfxReserva %>,
                                toolbar: { show: false }
                            },
                            colors: ['<%= Color %>'],
                            stroke: { width: 3, curve: 'smooth' },
                            dataLabels: { enabled: false },
                            xaxis: { categories: labels },
                            yaxis: { 
                                title: { text: 'Valor' },
                                min: 0,
                                max: <%= vMax %>
                            },
                            fill: {
                                type: 'solid',
                                opacity: 1
                            },
                            title: {
                                text: '<%= Titulo %>',
                                align: 'left',
                                style: {
                                    fontSize: '20px',
                                    fontWeight: 'bold',
                                    color: '#111111ff'
                                }
                            }
                        };

                        var chart = new ApexCharts(document.querySelector("#chart-line2"), options);
                        chart.render();
                    </script>
                 <%          
                end sub  

            ' Fin - Apex -------------------------------------------------------------   

            ' Tabs -------------------------------------------------------------------
    
                Sub framework_tabs(vector)
                    Dim k, lineas, txt

                    If Not IsArray(vector) Then Exit Sub
                    lineas = UBound(vector)

                    For k = 0 To lineas
                        txt = Trim(vector(k))

                        If InStr(txt, " ") > 0 Then
                            Response.Write "<div style='color:red;'>Error: '" & txt & "' contiene espacios.</div>"
                            Response.End
                        End If

                        If Len(txt) > 10 Then
                            Response.Write "<div style='color:red;'>Error: '" & txt & "' supera 10 caracteres.</div>"
                            Response.End
                        End If

                        txt = LCase(txt)
                        vector(k) = UCase(Left(txt,1)) & Mid(txt,2)
                    Next

                    Response.Write "<div class='fila-links'>"
                        For k = 0 To lineas
                            Response.Write "<a href='#' class='celda' onclick='verTab(" & (k + 1) & "); return false;'>" & vector(k) & "</a>"
                        Next
                    Response.Write "</div>"

                    Response.Write "<script>"
                        Response.Write "var secciones = [ " & vbCrLf
                            For k = 0 To lineas
                                If k > 0 Then Response.Write ","
                                Response.Write """tab_" & vector(k) & """" & vbCrLf
                            Next
                        Response.Write "];" & vbCrLf & vbCrLf

                        %>
                            function verTab(tab) {
                                // Ocultar todas las secciones
                                for (var i = 0; i < secciones.length; i++) {
                                    var sec = document.getElementById(secciones[i]);
                                    if (sec) sec.style.display = "none";
                                }

                                // Mostrar la sección activa
                                if (secciones[tab - 1]) {
                                    var activa = document.getElementById(secciones[tab - 1]);
                                    if (activa) activa.style.display = "block";
                                }

                                // Actualizar estado visual de los tabs
                                var tabs = document.querySelectorAll(".fila-links .celda");

                                for (var j = 0; j < tabs.length; j++) {
                                    tabs[j].classList.remove("activo");
                                }

                                if (tabs[tab - 1]) {
                                    tabs[tab - 1].classList.add("activo");
                                }
                            }                  
                        <%
                    response.write "</script>"
                End Sub

            ' Fin - Tabs -------------------------------------------------------------
        %>

        <style type="text/css">
            /* ===================================================
                        Framework 4 - styles.css
            =================================================== */

            /* Variables CSS */
                    
                :root {
                    --menu-barra-fondo:         <%= framework_mb_background() %>;
                    --menu-barra-font-color:    <%= framework_mb_font_color() %>;
                    --menu-listas:              <%= framework_menu_background_color() %>;
                    --menu-listas-oscuro:       hsl(from var(--menu-listas) h s calc(l - 4));
                    --menu-listas-profundo:     hsl(from var(--menu-listas) h s calc(l - 8));
                    --menu-listas-final:        hsl(from var(--menu-listas) h s calc(l - 12));
                    --menu-listas-final-cursor: hsl(from var(--menu-listas) h s calc(l - 16));
                    --menu-listas-font-color:   <%= framework_menu_font_Color() %>;

                    --verde:rgb(21, 140, 49);                       /* éxito */
                    --rojo:rgb(182, 68, 80);                        /* peligro */
                    --azul:rgb(56, 122, 197);                       /* info */
                    --naranja:rgb(215, 132, 63);                    /* advertencia */
                    --violeta:rgb(119, 68, 196);                    /* pregunta */    
                    --amarillo:rgb(219, 195, 41);                   /* comodín */    
                    --gris: rgba(153, 153, 153, 1);                 /* objeto */

                    --verde-claro: rgb(111, 197, 128);              /* éxito disabled */
                    --rojo-claro: rgb(211, 119, 127);               /* peligro disabled */
                    --azul-claro: rgb(123, 157, 221);               /* info disabled */
                    --naranja-claro: rgb(229, 168, 121);            /* advertencia  disabled */
                    --violeta-claro: rgb(159, 119, 216);            /* pregunta disabled */
                    --amarillo-claro: rgb(233, 223, 157);           /* comodín disabled */
                    --gris-claro: rgb(190, 190, 190);               /* objeto disabled */
                    
                    --btn-radius: 0.5rem;
                    --btn-size: 40px;   
                    
                    --field-background-color: rgb(236, 250, 237);

                    /* Tablas */

                        --tabla-fondo-head: rgb(244, 244, 244);
                        --tabla-fondo-body: rgb(255, 255, 255);
                        --tabla-fondo-foot: rgb(244, 244, 244);
                        
                        --tabla-borde: rgb(221, 221, 221);
                        --tabla-borde-body: rgb(221, 221, 221);
                        --tabla-borde-foot: rgb(221, 221, 221);
                        
                        --tabla-fondo-par-body: rgb(250, 250, 250);
                        --tabla-fondo-impar-body: rgb(255, 255, 255);

                        --tabla-texto: rgb(42, 42, 42);
                        --tabla-texto-head: rgb(34, 34, 34);
                        --tabla-texto-foot: rgb(34, 34, 34);
                        --tabla-hover-bg: rgb(233, 245, 255);

                    /* Fin de Tablas */                    
                } 

            /* Fin - Variables CSS */

            /* Reseteo de Clases y definifiones básicas */

                * {
                    margin: 0;
                    border: 0;
                    padding: 0;
                    font-size: 100%;
                }          

                body {
                    background-color: rgb(235, 235, 235); 
                    color: rgb(0, 0, 0); 
                    font: 14px Arial, sans-serif;        
                }

                body, html { 
                    margin: 0; 
                    padding: 0; 
                    height: 100%; 
                    overflow: hidden; 
                }

                a {
                    text-decoration: none;
                    outline: none;
                }

                input, select {
                    border-radius: var(--btn-radius);
                }

                .roundTop {
                    border-radius: var(--btn-radius) var(--btn-radius);
                }

                table {
                    border-collapse: collapse;
                    width: 100%;
                    border-spacing: 0;
                }

                table, tr, td, th {
                    border: none;
                    padding: 5px;
                }

                tr:not(:last-child) {
                    border-bottom: 1px solid #ccc; /* se puede ajustar, dependiendo del contexto */
                } 
                
                .no-ver {
                    display: none   !important;
                    width:   0      !important;
                    height:  0      !important;
                    padding: 0      !important;
                    margin:  0      !important;
                    border:  none   !important;
                }   
                
                .noborder {
                    border: none !important;
                }  
                
                .clickable {
                    cursor: pointer;
                }                   
                  
            /* Fin - Reseteo de Clases y definifiones básicas */

            /* Tablas */

                /*  Clases Base Para Tablas */

                    .tabla-wrapper {
                        width: 100%;
                        overflow-y: auto;
                        overflow-x: hidden;
                        -webkit-overflow-scrolling: touch;

                        border-radius: 12px;
                        box-shadow: 0 0 6px var(--tabla-borde);
                        background-color: var(--tabla-fondo-body);
                    }                

                    .tabla {
                        width: 100%;
                        border-collapse: collapse;
                        border-spacing: 0;
                        font-family: system-ui, sans-serif;
                        font-size: 0.95rem;
                        color: var(--tabla-texto);
                    }

                    .tabla thead {
                        background-color: var(--tabla-fondo-head);
                        color: var(--tabla-texto-head);
                        text-transform: uppercase;
                        font-weight: bold;
                    }

                    .tabla thead tr th:first-child {
                        border-top-left-radius: 12px;
                    }

                    .tabla thead tr th:last-child {
                        border-top-right-radius: 12px;
                    }           

                    .tabla tfoot {
                        background-color: var(--tabla-fondo-foot);
                        color: var(--tabla-texto-foot);
                        text-transform: uppercase;
                        font-weight: bold;
                    }                

                    .tabla tfoot tr td:first-child {
                        border-bottom-left-radius: 12px;
                    }

                    .tabla tfoot tr td:last-child {
                        border-bottom-right-radius: 12px;
                    }
                    
                    .tabla th,
                    .tabla td {
                        padding: 0.75rem 1rem;
                        border: 1px solid var(--tabla-borde-body);
                        text-align: left;
                        vertical-align: middle !important;
                    }

                    .tabla thead th.sticky {
                        position: sticky;
                        top: 0;
                        background-color: var(--tabla-fondo-head);
                        z-index: 10;
                    }

                    .tabla tfoot td.sticky {
                        position: sticky;
                        bottom: 0;
                        background-color: var(--tabla-fondo-foot);
                        border: 1px solid var(--tabla-borde-foot);
                        z-index: 10;
                    }

                    .tabla tbody tr:nth-child(even) {
                        background-color: var(--tabla-fondo-par-body);
                    }

                    .tabla tbody tr:nth-child(odd) {
                        background-color: var(--tabla-fondo-impar-body);
                    }

                    .tabla tbody tr:hover {
                        background-color: var(--tabla-hover-bg);
                        transition: background-color 0.2s ease-in-out;
                    }

                /*  Fin - Clases Base Para Tablas */      
                
                /*  "Skins" Para Tablas */

                    .tabla-carbon, .tabla-gris {
                        --tabla-fondo-head: rgba(91, 91, 91, 1);
                        --tabla-fondo-body: rgba(234, 234, 234, 1);
                        --tabla-fondo-foot: rgba(91, 91, 91, 1);
                        
                        --tabla-borde: rgba(149, 149, 149, 1);
                        --tabla-borde-body: rgba(222, 222, 222, 1);
                        --tabla-borde-foot: rgba(149, 149, 149, 1);
                        
                        --tabla-fondo-par-body: rgba(230, 230, 230, 1);
                        --tabla-fondo-impar-body: rgb(255, 255, 255);

                        --tabla-texto: rgba(0, 0, 0, 1);
                        --tabla-texto-head: rgba(255, 255, 255, 1);
                        --tabla-texto-foot: rgba(255, 255, 255, 1);
                        --tabla-hover-bg: rgba(240, 240, 240, 1);      
                    }

                    .tabla-blue, .tabla-azul {
                        --tabla-fondo-head: rgba(13, 71, 99, 1);
                        --tabla-fondo-body: rgba(229, 244, 252, 1);
                        --tabla-fondo-foot: rgba(13, 71, 99, 1);
                        
                        --tabla-borde: rgba(96, 123, 148, 1);
                        --tabla-borde-body: rgba(209, 231, 250, 1);
                        --tabla-borde-foot: rgba(15, 58, 95, 1);
                        
                        --tabla-fondo-par-body: rgba(240, 250, 255, 1);
                        --tabla-fondo-impar-body: rgba(223, 245, 255, 1);

                        --tabla-texto: rgba(0, 0, 0, 1);
                        --tabla-texto-head: rgba(255, 255, 255, 1);
                        --tabla-texto-foot: rgba(255, 255, 255, 1);
                        --tabla-hover-bg: rgba(203, 226, 245, 1);                         
                    }    
                    
                    .tabla-green, .tabla-verde {
                        --tabla-fondo-head: rgba(15, 82, 10, 1);
                        --tabla-fondo-body: rgba(217, 244, 223, 1);
                        --tabla-fondo-foot: rgba(15, 82, 10, 1);
                        
                        --tabla-borde: rgba(44, 70, 28, 1);
                        --tabla-borde-body: rgba(205, 219, 197, 1);
                        --tabla-borde-foot: rgba(44, 70, 28, 1);
                        
                        --tabla-fondo-par-body: #F3FAEE;
                        --tabla-fondo-impar-body: #e4efddff;

                        --tabla-texto: rgba(0, 0, 0, 1);
                        --tabla-texto-head: rgba(255, 255, 255, 1);
                        --tabla-texto-foot: rgba(255, 255, 255, 1);
                        --tabla-hover-bg: rgba(223, 246, 214, 1);     
                    }  
                    
                    .tabla-violet, .tabla-violeta {
                        --tabla-fondo-head: rgba(57, 10, 82, 1);
                        --tabla-fondo-body: rgba(235, 217, 244, 1);
                        --tabla-fondo-foot: rgba(60, 10, 82, 1);
                        
                        --tabla-borde: rgba(60, 28, 70, 1);
                        --tabla-borde-body: rgba(215, 197, 219, 1);
                        --tabla-borde-foot: rgba(63, 28, 70, 1);
                        
                        --tabla-fondo-par-body: #f5eefaff;
                        --tabla-fondo-impar-body: #e9ddefff;

                        --tabla-texto: rgba(0, 0, 0, 1);
                        --tabla-texto-head: rgba(255, 255, 255, 1);
                        --tabla-texto-foot: rgba(255, 255, 255, 1);
                        --tabla-hover-bg: rgba(232, 214, 246, 1);     
                    } 
                    
                    .tabla-red, .tabla-roja {
                        --tabla-fondo-head: rgba(70, 28, 28, 1);
                        --tabla-fondo-body: rgba(244, 217, 217, 1);
                        --tabla-fondo-foot: rgba(70, 28, 28, 1);
                        
                        --tabla-borde: rgba(84, 36, 36, 1);
                        --tabla-borde-body: rgba(219, 197, 197, 1);
                        --tabla-borde-foot: rgba(70, 28, 28, 1);
                        
                        --tabla-fondo-par-body: #faeeeeff;
                        --tabla-fondo-impar-body: #efddddff;

                        --tabla-texto: rgba(0, 0, 0, 1);
                        --tabla-texto-head: rgba(255, 255, 255, 1);
                        --tabla-texto-foot: rgba(255, 255, 255, 1);
                        --tabla-hover-bg: rgba(246, 214, 214, 1);                           
                    }          
                    
                /*  Fin - Skins Para Tablas */                    
                
                /* Tabla neutra / transparente */
                
                    .tabla-transparente,
                    .tabla-transparente table,
                    .tabla-transparente tr,
                    .tabla-transparente td,
                    .tabla-transparente th {
                        background: transparent !important;
                        border: none !important;
                        box-shadow: none !important;
                    }  
                                
                /* Fin: Tabla Transparente */
                
                /*  Modificadores de Lineas (para casos especiales) */

                    .tr-verde { background-color: #f1ffedff !important; }
                    .tr-verde:hover { background-color: #ddffd2ff !important; }         
                    
                    .tr-rojo { background-color: #fbebebff !important; }
                    .tr-rojo:hover { background-color: #fde3e3ff !important; }

                    .tr-azul { background-color: #e8f0ffff !important; }
                    .tr-azul:hover { background-color: #dde9ffff !important; }

                    .tr-morado { background-color: #f3ebffff !important; }            
                    .tr-morado:hover { background-color: #ecdeffff !important; }  
                    
                /* Fin - Modificadores */                

            /* Fin - Tablas mediante <div> */

            /* Menues y Navegación */
                
                nav { background-color: var(--menu-barra-fondo); }

                .nav-menu, .nav-menu ul,
                .nav-menu li, .nav-menu a {
                    margin: 0;
                    padding: 0;
                    line-height: normal;
                    list-style: none;
                    display: block;
                    position: relative;
                }

                .nav-menu ul {
                    opacity: 0;
                    position: absolute;
                    top: 100%;
                    left: -9999px;
                    z-index: 999;
                    -webkit-transition: opacity .3s;
                    transition: opacity .3s;
                }

                .nav-menu li:hover>ul {
                    left: 0;
                    opacity: 1;
                    z-index: 1000;
                }

                .nav-menu ul li:hover>ul {
                    top: 0;
                    left: 100%;
                }

                .nav-menu li {
                    cursor: default;
                    float: left;
                    white-space: nowrap;
                }

                .nav-menu ul li {
                    float: none;
                }

                .nav-menu ul {
                    min-width: 12em;
                    -webkit-box-shadow: 2px 2px 2px rgba(0, 0, 0, 0.4), 0 0 0 0 rgba(0, 0, 0, 0.19);
                    box-shadow: 0 6px 6px 0 rgba(0, 0, 0, 0.4), 0 0 0 0 rgba(0, 0, 0, 0.19);
                }

                .nav-menu a {
                    padding: 0 10px;
                    color: var(--menu-barra-font-color);
                    font-weight: normal;
                    line-height: 40px;
                    text-decoration: none;
                }

                .nav-menu>li>.nav-active {
                    /*
                        COLOR DE LA BARRA DE MENU
                    */
                    background-color: rgb(255, 255, 255); 
                    color: rgb(33, 33, 33);  
                }

                .nav-menu a:hover,
                .nav-menu a:focus,
                .nav-menu li:hover a {
                    background: var(--menu-listas-oscuro);
                    color: <%= framework_menu_font_Color() %>;
                }

                .nav-menu li li a,
                .nav-menu li:hover li a {
                    padding: 8px 10px;
                    background: var(--menu-listas);
                    color: <%= framework_menu_font_Color() %>;
                    line-height: normal;
                }

                .nav-menu li:hover li a:hover,
                .nav-menu li:hover li a:focus,
                .nav-menu li:hover li:hover a {
                    background: var(--menu-listas-oscuro);
                }

                .nav-menu li:hover li:hover li a {
                    background: var(--menu-listas-oscuro);
                }

                .nav-menu li:hover li:hover li a:hover,
                .nav-menu li:hover li:hover li a:focus,
                .nav-menu li:hover li:hover li:hover a {
                    background: var(--menu-listas-final);
                }

                .nav-menu li:hover li:hover li:hover li a {
                    background: var(--menu-listas-final);
                }

                .nav-menu li:hover li:hover li:hover li a:hover,
                .nav-menu li:hover li:hover li:hover li a:focus {
                    background: var(--menu-listas-final-cursor);
                }    

                .sombra {
                    box-shadow: 0 6px 6px 0 rgba(0, 0, 0, <%= framework_menu_drop_shadow() %>), 0 0 0 0 rgba(0, 0, 0, 0.19);
                }  

            /* Fin - Menúes y navegación */

            /* Definición de Snippets para el Modo Escritorio */
            
                .winBody {
                    position: absolute;
                    display: block;
                    background-color: #f1f1f1;
                    font-family: 'Ruda', sans-serif;
                    font-size: 14px;                
                    text-align: left;
                    border: 1px solid #6f6f6f;
                    border-radius: 10px;
                    
                    opacity: <%= Opacidad() %>;
                    box-shadow: 0 6px 6px 0 rgba(0, 0, 0, 0.4), 0 0 0 0 rgba(0, 0, 0, 0.19);
                    z-index: 5;   
                    user-select: none;       
                }

                .winHeader {
                    border-radius: 10px 10px 0px 0px;
                    font-family: 'Ruda', sans-serif;
                    font-size: 12px;
                    padding: 5px;
                    cursor: move;
                    background-color: rgb(25, 33, 64);
                    color: #fff;
                    opacity: <%= Opacidad() %>;                    
                    z-index: 6;          
                }

            /* Fin - Snippets */

            /*  Helpers del Framework */

                .framework_mb {
                    display: flex;
                    justify-content: space-between;
                    width: 100%; /* Se ajusta al tamaño del contenedor/pantalla */
                    box-sizing: border-box; /* Asegura que padding no altere el tamaño total */
                }
                
                .framework_menu {
                    text-align: left; /* Opcional, ya que es el predeterminado */
                }
                
                .framework_shortcuts {  
                    text-align: right; /* Opcional, pero asegura el estilo */                
                }   

            /* Fin - Helpers */

            /* Fonts */

                @font-face { font-family: Allura; src: url('/core/fonts/allura.otf'); }            
                @font-face { font-family: 'Arial Narrow'; src: url('/core/fonts/arialnarrow.otf'); }
                @font-face { font-family: Beyond; src: url('/core/fonts/beyond.otf'); }
                @font-face { font-family: Caviar; src: url('/core/fonts/caviar.otf'); }
                @font-face { font-family: Domotika; src: url('/core/fonts/domotika.otf'); }
                @font-face { font-family: Dosis; src: url('/core/fonts/dosis.otf'); }
                @font-face { font-family: Harting; src: url('/core/fonts/harting.otf'); }
                @font-face { font-family: Kaushan; src: url('/core/fonts/kaushan.otf'); }
                @font-face { font-family: Lobster; src: url('/core/fonts/lobster.otf'); }
                @font-face { font-family: Ruda; src: url('/core/fonts/ruda.otf'); }
                @font-face { font-family: 'Ruda Bold'; src: url('/core/fonts/ruda_bold.otf'); }
                @font-face { font-family: Aptos; src: url('/core/fonts/aptos.otf'); }            
                @font-face { font-family: GameStation; src: url('/core/fonts/gamestation.otf'); }
                @font-face { font-family: Mark Script; src: url('/core/fonts/mark.otf'); }
                @font-face { font-family: Season; src: url('/core/fonts/season.otf'); }
                @font-face { font-family: Titillium; src: url('/core/fonts/titillium.otf'); }

            /* Fin - Fonts */

            /* Botones y Tamaños */

                .form-btn.verde    { --btn-color: var(--verde); }
                .form-btn.rojo     { --btn-color: var(--rojo); }
                .form-btn.azul     { --btn-color: var(--azul); }
                .form-btn.naranja  { --btn-color: var(--naranja); }
                .form-btn.violeta  { --btn-color: var(--violeta); }
                .form-btn.amarillo { --btn-color: var(--amarillo); }
                .form-btn.gris     { --btn-color: var(--gris); }

                .form-btn.verde.disabled    { --btn-color-disabled: var(--verde-claro); }
                .form-btn.rojo.disabled     { --btn-color-disabled: var(--rojo-claro); }
                .form-btn.azul.disabled     { --btn-color-disabled: var(--azul-claro); }
                .form-btn.naranja.disabled  { --btn-color-disabled: var(--naranja-claro); }
                .form-btn.violeta.disabled  { --btn-color-disabled: var(--violeta-claro); }            
                .form-btn.amarillo.disabled { --btn-color-disabled: var(--amarillo-claro); }     
                .form-btn.gris.disabled     { --btn-color-disabled: var(--gris-claro); }     

                /* Clase base, ya adaptada */

                .form-btn {
                    display: inline-flex;
                    align-items: center;
                    justify-content: center;
                    padding: 0 0.75rem;
                    font-family: 'Ruda';
                    font-size: 1rem;
                    line-height: 1;
                    border: 1px solid var(--bs-border-color);
                    border-right: 0;
                    border-radius: var(--btn-radius);
                    color: rgb(255, 255, 255);
                    background-color: var(--btn-color, transparent);
                    width: var(--btn-size);
                    height: var(--btn-size);
                }

                .form-btn.disabled {
                    pointer-events: none;
                    opacity: 0.8;
                    background-color: var(--btn-color-disabled, #ccc);
                    border-color: var(--btn-color-disabled, #ccc);
                }  

                .form-btn.small  { min-width:  60px; font-size: 1rem; color: rgb(255, 255, 255); }
                .form-btn.normal { min-width: 100px; font-size: 1rem; color: rgb(255, 255, 255); }
                .form-btn.large  { min-width: 200px; font-size: 1rem; color: rgb(255, 255, 255); }
                .form-btn.xl     { min-width: 300px; font-size: 1rem; color: rgb(255, 255, 255); }
                .form-btn.xxl    { min-width: 400px; font-size: 1rem; color: rgb(255, 255, 255); }     

            /* Fi - Botones y Tamaños */

            /*  Formularios */

                .main {
                    max-width: 90%;
                    margin: 0.5rem auto;
                    padding: 1.5rem;
                    background-color: #fff;
                    border-radius: 0.5rem;
                    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.4);
                    display: flex;
                    flex-direction: column;
                    gap: 1.5rem;
                    font-family: sans-serif;
                    transition: box-shadow 0.3s ease;
                }

                .main-scroll {
                    max-height: calc(100vh - 220px); /* margenes incluidos */
                    overflow-y: auto;
                    overflow-x: hidden;
                    padding-right: 1rem;
                }

                .main:hover,
                .main:focus-within {
                    box-shadow: 0 4px 10px rgba(0, 0, 0, 0.6);
                }

                .line {
                    display: flex;
                    flex-direction: row;
                    gap: 1rem;
                }

                .label {
                    color: rgb(0, 0, 0);
                    text-align: left;
                    padding: 0.3rem 0.4rem;
                    font-family: 'Ruda Bold', sans-serif;
                    font-size: 1rem;

                    display: flex;
                    align-items: center; 
                }

                .field {
                    padding: 0.3rem 0.4rem;
                    border: 1px solid #ccc;
                    border-radius: 0.3rem;
                    font-family: 'Ruda', sans-serif;
                    font-size: 1rem;
                    color: rgb(25, 25, 25);                
                    background-color: var(--field-background-color);
                    box-sizing: border-box;
                    resize: vertical;
                }

                .field.frame {
                    background-color: transparent !important;
                    border: none !important;
                    box-shadow: none !important;
                    padding: 0.3rem 0.4rem;
                    width: 100%;
                    font-family: inherit;
                    font-size: inherit;
                    color: inherit;
                    transition: box-shadow 0.2s ease;
                }
               
                .field.frame:focus {
                    box-shadow: 0 0 0 2px rgba(0, 0, 0, 0.3);   /* Glow visible */
                    border-radius: 4px;
                }                

                .no-field {
                    padding: 0.3rem 0.4rem;
                    border: 1px solid #ccc;
                    border-radius: 0.3rem;
                    font-family: 'Ruda', sans-serif;
                    font-size: 1rem;
                    color: rgb(25, 25, 25);                
                    background-color: var(--field-background-color);
                    box-sizing: border-box;
                    resize: vertical;
                }                
                
                .section {
                    display: block      !important;
                    align-items: unset  !important;
                    border: 1px solid rgba(230, 230, 230, 1); 
                    border-radius: var(--btn-radius) var(--btn-radius);
                    padding: 1rem;
                }

                .doble-wrapper{
                    display:flex;
                    gap:8px;
                    flex-wrap:wrap;
                }

                /* Sub-Clases Especiales */

                    /* 
                        El Label del <div class="line"> 
                        aparecerá en la parte superior 
                        (vertical-align: top;)
                    */

                    .line.label-top {
                        align-items: flex-start;
                    }

                    .line.label-top > label {
                        padding-top: 6px; /* ajustable según tu estética */
                    }

                    /* 
                        Si dentro de <div class='line'> se añade otro div que
                        contendrá campos (class='field'), deben estar contenidos
                        dentro de un <div class='line line-group'> para mantener
                        la coherencia
                    */                    

                    .line-group .line {
                        margin-bottom: 12px;
                        margin-right: 20px;
                    }

                    .line-group .line:last-child {
                        margin-bottom: 0;
                        margin-right: 20px;
                    }                      

                /* Fin - Clases Especiales */                

                /*
                    Al definir los tamaños como width y no min-width estamos creando,
                    de forma disimulada, un esquema de columnas fijas
                */

                .label.tiny   { width:  75px; }
                .label.small  { width: 100px; }
                .label.normal { width: 150px; }
                .label.large  { width: 200px; }
                .label.xl     { width: 250px; }
                .label.xxl    { width: 300px; }
                .label.xxl2   { width: 500px; }
                .label.full   { width: 100% ; }

                .field.tiny   { width: 125px; }
                .field.small  { width: 200px; }
                .field.normal { width: 300px; }
                .field.large  { width: 500px; }
                .field.xl     { width: 600px; }
                .field.xxl    { width: 700px; }
                .field.full   { width: 100% ; } /* Esto es para elementos dentro de un contenedor */

                @media (max-width: 1000px) {
                    .line {
                        flex-direction: column;
                        align-items: stretch;
                        gap: 0; 
                    }

                    .label {
                        width: 100%      !important;
                        text-align: left !important;
                        padding-right: 0;
                    }

                    .field {
                        width: 100%     !important;
                        min-width: 0    !important;
                        max-width: 100% !important; 
                        flex: 1         !important; 
                    }

                    .seccion {
                        width: 95%         !important;
                        text-align: center  !important;

                        display: block      !important;
                        align-items: unset  !important;
                        border: 1px solid rgba(230, 230, 230, 1); 
                        border-radius: var(--btn-radius) var(--btn-radius);
                        padding: 1rem;
                    }                    
                }                 

            /* Fin - Formularios */

            /* SWAL CSS */

                .swal-text { 
                    font-size: 16px;
                    font-family: monospace, monospace;
                    padding: 0 10px;
                    font-weight: 400;
                    color: rgba(0,0,0,.64); 
                    max-width: 100%;
                }

                .swal-modal {
                    width: 600px;
                    text-align: center;
                    border-radius: 5px;
                }  
                
                .swal2-popup-perfil {
                    font-size: 14px;
                    padding: 10px;
                    width: 280px;
                    background: #f2f2f2;
                    color: #333;
                }        
                        
            /* Fin - Swal */  

            /* Pop Up de Notificaciones Globales */
                /* Modal base */
                    .modal-f4 {
                        position: fixed;
                        top: 0;
                        left: 0;
                        width: 100vw;
                        height: 100vh;
                        background: rgba(40, 40, 40, 0.5);
                        display: flex;
                        justify-content: center;
                        align-items: center;
                        z-index: 9999;
                    }

                /* Oculto por defecto */
                    .modal-f4.oculto {
                        display: none;
                    }

                /* Contenedor interno */
                    .modal-contenido {
                        width: 600px;
                        max-height: 80vh;
                        background: #ffffff;
                        border-radius: 12px;
                        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.25);
                        overflow: hidden;
                        display: flex;
                        flex-direction: column;
                        animation: entradaModal 0.3s ease-out;
                        font-family: 'Segoe UI', Tahoma, sans-serif;
                    }

                /* Animación de entrada */
                    @keyframes entradaModal {
                        from { transform: translateY(-20px); opacity: 0; }
                        to   { transform: translateY(0); opacity: 1; }
                    }

                /* Encabezado */
                    .modal-header {
                        background-color: #f4f4f4;
                        padding: 12px 16px;
                        display: flex;
                        justify-content: space-between;
                        align-items: center;
                        border-bottom: 1px solid #ddd;
                    }

                    .modal-titulo {
                        font-weight: bold;
                        font-size: 1.1rem;
                        color: #333;
                    }

                    .modal-cerrar {
                        cursor: pointer;
                        font-size: 1.4rem;
                        color: #666;
                        transition: color 0.2s;
                    }

                    .modal-cerrar:hover {
                        color: #000;
                    }

                /* Lista de notificaciones */
                    .modal-lista {
                        padding: 12px 16px;
                        overflow-y: auto;
                        flex-grow: 1;
                    }

                /* Cada item */
                    .modal-item {
                        padding: 12px;
                        margin-bottom: 10px;
                        border-left: 5px solid #ccc;
                        background-color: #f5eded;
                        border-radius: 8px;
                        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                    }

                    .modal-item .item-titulo {
                        font-weight: bold;
                        font-size: 1rem;
                        margin-bottom: 4px;
                        color: #333;
                    }

                    .modal-item .item-mensaje {
                        font-size: 0.95rem;
                        color: #555;
                        margin-bottom: 6px;
                    }

                    .modal-item .item-fecha {
                        font-size: 0.8rem;
                        color: #999;
                        text-align: right;
                    }

                /* Estilos por tipo */
                    .tipo-A {
                        border-left-color: #007bff; /* Azul */
                        background-color: #cce0f5;
                    }

                    .tipo-E {
                        border-left-color: #dc3545; /* Rojo */
                        background-color: #f1cece;
                    }

                    .tipo-P {
                        border-left-color: #ffc107; /* Amarillo */
                        background-color: #eee4be;
                    }

                /* Mensajes especiales */
                    .modal-vacio,
                    .modal-error {
                        text-align: center;
                        color: #999;
                        padding: 20px;
                        font-size: 0.95rem;
                    }            
            /* Fin - Pop Up de Notificaciones Globales */

            /* Objetos Plus! */

                /* Objetos Plus - Modo Lista (Vista por Defecto) ------------*/
                
                    .objeto-plus {
                        position: absolute;

                        background-color: var(--menu-listas);
                        color: var(--menu-listas-font-color);
                        font-family: <%= framework_mb_font() %>;
                        font-size: <%= framework_menu_font_size() %>px;

                        border-radius: 6px;
                        box-shadow: 0 6px 6px 0 rgba(0, 0, 0, <%= framework_menu_drop_shadow() %>),
                                    0 0 0 0 rgba(0, 0, 0, 0.19);

                        padding: 8px 0;
                        min-width: 140px;
                        max-height: 300px;
                        overflow-y: auto;
                        z-index: 1000;
                    }

                    .objeto-plus {
                        display: none;
                    }

                    .objeto-plus.plus-visible {
                        display: block;
                    }

                    .objeto-plus.plus-visible.plus-modo-botones {
                        display: flex;
                    }                

                    .objeto-plus .plus-item {
                        display: flex;
                        align-items: center;
                        gap: 8px;                 /* Si hay <i>, separa. Si no, no estorba */
                        padding: 8px 12px;
                        cursor: pointer;
                        user-select: none;
                    }

                    .objeto-plus .plus-item:hover { background-color: var(--menu-listas-oscuro); }
                    .objeto-plus .plus-item:active { background-color: var(--menu-listas-profundo); }
                    .objeto-plus .plus-item i { flex-shrink: 0; /* Evita que el icono se deforme */ }
                    .objeto-plus .plus-item span { flex: 1; /* El texto ocupa el espacio disponible */ }

                /*-----------------------------------------------------------*/

                /* Objetos Plus - Modo Botones (Compacto) -------------------*/

                    .objeto-plus.plus-visible.plus-modo-botones {
                        display: inline-flex;
                        flex-wrap: wrap;
                        gap: 4px;
                        padding: 6px;

                        min-width: 0;
                        width: calc((56px * 2) + 4px + 12px);
                    }

                    /* Centrar el bloque si está dentro de un contenedor más grande */
                    .objeto-plus.plus-modo-botones-wrapper {
                        text-align: center;
                    }

                    .objeto-plus.plus-modo-botones .plus-item {
                        width: 54px;
                        min-height: 44px;

                        display: flex;
                        flex-direction: column;
                        align-items: center;
                        justify-content: center;

                        text-align: center;
                        padding: 4px;

                        border-radius: 5px;
                    }

                    .objeto-plus.plus-modo-botones .plus-item i {
                        font-size: 15px;
                        margin-bottom: 2px;
                    }

                    .objeto-plus.plus-modo-botones .plus-item span {
                        font-size: 10px;
                        line-height: 1.05;
                    }
                    
                /*-----------------------------------------------------------*/

                /* Objetos Plus - Modo Cinta (Ultra Compacto) ---------------*/

                    .objeto-plus.plus-visible.plus-modo-cinta {
                        display: flex !important;
                        flex-wrap: wrap;
                        gap: 4px;
                        padding: 4px;

                        width: fit-content;
                    }

                    .objeto-plus.plus-modo-cinta .plus-item {
                        width: 42px;           /* Más pequeño */
                        height: 42px;          /* Cuadrado perfecto */

                        display: flex           !important;
                        align-items: center     !important;
                        justify-content: center !important;

                        padding: 0;
                        border-radius: 6px;
                    }

                    /* Ocultamos el texto */
                    .objeto-plus.plus-modo-cinta .plus-item span {
                        display: none;
                    }

                    /* Ícono ligeramente más grande ahora que es protagonista */
                    .objeto-plus.plus-modo-cinta .plus-item i {
                        font-size: 18px;
                        margin: 0;
                    }
                /*-----------------------------------------------------------*/

            /* Fin - Objetos Plus! */

            /* Tabs */ 
                .fila-links {
                    font-family: Ruda;
                    font-size: 16px;
                    display: flex;
                    width: 95%;
                    margin: 0 auto;
                    border-radius: 6px;
                    overflow: hidden;
                    border: none;
                }
            
                .fila-links .celda.activo {
                    background: rgb(118, 54, 153);
                    color: #ffffff;
                    font-weight: 600;
                    
                }

                .fila-links .celda {
                    flex: 1;
                    padding: 10px;
                    text-align: center;
                    text-decoration: none;
                    background: rgb(88, 36, 117);
                    color: #ffffff;
                    font-size: 14px;
                    font-weight: 500;
                    padding: 10px 14px;
                    text-transform: none;                
                    border: none;
                }

                .fila-links .celda:hover {
                    background: rgb(143, 74, 180); /* micro contraste, nada suave */
                } 
            /* Fin - Tabs */             
        </style>