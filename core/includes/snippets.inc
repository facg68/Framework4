<style>
    /* Botones y Tamaños */

        .snippet-btn.verde    { --btn-color: var(--verde); }
        .snippet-btn.rojo     { --btn-color: var(--rojo); }
        .snippet-btn.azul     { --btn-color: var(--azul); }
        .snippet-btn.naranja  { --btn-color: var(--naranja); }
        .snippet-btn.violeta  { --btn-color: var(--violeta); }
        .snippet-btn.amarillo { --btn-color: var(--amarillo); }
        .snippet-btn.gris     { --btn-color: var(--gris); }

        .snippet-btn.verde.disabled    { --btn-color-disabled: var(--verde-claro); }
        .snippet-btn.rojo.disabled     { --btn-color-disabled: var(--rojo-claro); }
        .snippet-btn.azul.disabled     { --btn-color-disabled: var(--azul-claro); }
        .snippet-btn.naranja.disabled  { --btn-color-disabled: var(--naranja-claro); }
        .snippet-btn.violeta.disabled  { --btn-color-disabled: var(--violeta-claro); }            
        .snippet-btn.amarillo.disabled { --btn-color-disabled: var(--amarillo-claro); }     
        .snippet-btn.gris.disabled     { --btn-color-disabled: var(--gris-claro); }     

        /* Clase base, ya adaptada */

        .snippet-btn {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            padding: 4px;
            font-family: 'Ruda';
            line-height: 1;
            border: 1px solid var(--bs-border-color);
            border-right: 0;
            border-radius: 0.25rem;;
            color: rgb(255, 255, 255);
            background-color: var(--btn-color, transparent);
            width: var(--btn-size);
        }

        .snippet-btn.disabled {
            pointer-events: none;
            opacity: 0.8;
            background-color: var(--btn-color-disabled, #ccc);
            border-color: var(--btn-color-disabled, #ccc);
        }  

        .snippet-btn.small  { min-width:  60px; color: rgb(255, 255, 255); }
        .snippet-btn.normal { min-width: 100px; color: rgb(255, 255, 255); }
        .snippet-btn.large  { min-width: 200px; color: rgb(255, 255, 255); }
        .snippet-btn.xl     { min-width: 300px; color: rgb(255, 255, 255); }
        .snippet-btn.xxl    { min-width: 400px; color: rgb(255, 255, 255); }     

    /* Fi - Botones y Tamaños */
</style>        

<%
    Function framework_param(param, default)
        framework_param = ParametroUsuario("seguridad", param)    

        if isnull(framework_param) or (framework_param = "") then
            framework_param =  DefaultParametro("seguridad", param)

            if isnull(framework_param) or (framework_param = "") then
                framework_param = default
            end if
        end if                  
    end Function

    Function framework_menu_font_size()
        framework_menu_font_size = framework_param("title_area_font_size", "14px")
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

    sub apexColumns(Titulo, NombreDivChart, SQL_Query, Codigo, Valor, color, gfxReserva, ancho)
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
                            gLabel = gLabel & "['" & gfxTable(0) & "'"
                            gData1 = gData1 & "[" & gfxTable(1) 
                        else
                            gLabel = gLabel & ", '" & gfxTable(0) & "'"
                            gData1 = gData1 & ", " & gfxTable(1)                                     
                        end if

                        gfxTable.MoveNext
                    loop until gfxTable.eof

                    gLabel = gLabel & "]"
                    gData1 = gData1 & "]"
                end if
            gfxTable.close: set gfxTable = nothing
        gfxConn.close: set gfxConn = nothing

        response.write "<div id='" & NombreDivChart & "' style='padding: 0px; width: " & ancho & "px;'></div>"

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

    sub apexDonut(Titulo, NombreDivChart, SQL_Query, Codigo, Valor, gfxReserva, ancho)
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
                            gLabel = gLabel & "['" & gfxTable(0) & "'"
                            gData1 = gData1 & "[" & gfxTable(1) 
                        else
                            gLabel = gLabel & ", '" & gfxTable(0) & "'"
                            gData1 = gData1 & ", " & gfxTable(1)                                     
                        end if

                        gfxTable.MoveNext
                    loop until gfxTable.eof

                    gLabel = gLabel & "]"
                    gData1 = gData1 & "]"
                end if
            gfxTable.close: set gfxTable = nothing
        gfxConn.close: set gfxConn = nothing

        response.write "<div id='" & NombreDivChart & "' style='padding: 0px; width: " & ancho & "px;'></div>"
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

    sub apexPanoramic(Titulo, SQL_Query, Codigo, Valor, Color, gfxReserva, ancho)
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
                            gLabel = gLabel & "[" & gfxTable(0) 
                            gData1 = gData1 & "[" & gfxTable(1) 
                        else
                            gLabel = gLabel & ", " & gfxTable(0) 
                            gData1 = gData1 & ", " & gfxTable(1)                                     
                        end if

                        if cDbl(gfxTable(1) ) > vMax then vMax = cDbl(gfxTable(1) )

                        gfxTable.MoveNext
                    loop until gfxTable.eof

                    gLabel = gLabel & "]"
                    gData1 = gData1 & "]"
                end if
            gfxTable.close: set gfxTable = nothing
        gfxConn.close: set gfxConn = nothing      
        %>

        <div>
            <div id="chart-line2" style="width: <%= ancho %>px;"></div>
            <div id="chart-line"  style="width: <%= ancho %>px;"></div>
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

    sub apexArea(Titulo, SQL_Query, Codigo, Valor, Color, gfxReserva, ancho)
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
                            gLabel = gLabel & "[" & gfxTable(0) 
                            gData1 = gData1 & "[" & gfxTable(1) 
                        else
                            gLabel = gLabel & ", " & gfxTable(0) 
                            gData1 = gData1 & ", " & gfxTable(1)                                     
                        end if

                        if cDbl(gfxTable(1) ) > vMax then vMax = cDbl(gfxTable(1) )

                        gfxTable.MoveNext
                    loop until gfxTable.eof

                    gLabel = gLabel & "]"
                    gData1 = gData1 & "]"
                end if
            gfxTable.close: set gfxTable = nothing
        gfxConn.close: set gfxConn = nothing    

        %>
            <div id="chart-line2" style="width: <%= ancho %>px;"></div>

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
%>