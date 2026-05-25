        <% 
            if (thisSystem <> "") AND (thisProcess <> "") then
                Bitacora 
            end if

            title_background_color = framework_param("title_area_color", "#000000")
            title_area_font_color  = framework_param("title_area_font_color", "#FFFFFF")       
            logo_foreground_color  = framework_param("title_area_logo_color", "#FFFFFF")       
            icon_foreground_color  = framework_param("title_area_icons_color", "#FFFFFF")       

            styleMenuBar = ""

            '
            ' Fuente del Menu y de las listas de opciones
            '
            styleMenuBar = styleMenuBar & "font-family: " & framework_mb_font() & "; "
            styleMenuBar = styleMenuBar & "font-size: " & framework_menu_font_size() & "px; "          
            styleMenuBar = styleMenuBar & "color: " & framework_mb_font_color() & ";"  
        %>

        <script>
            var pageReserva = 190; // Reserva "por defecto" (Título + Filtros + espacio bajo el main)
            
            var menuReserva = <% 
                if (NoVerEncabezado() = 1) Or (thisNoMenu = True) then
                    response.write "21"
                else
                    response.write "77"
                end if
            %>;

            const getVal = id => document.getElementById(id).value;
            const setVal = (id, val) => document.getElementById(id).value = val;

            let mb_Maximizar = () => {};
            let mb_Minimizar = () => {};               
                    
            var core_desktop_maxWP = <%= Request.Cookies("max_WP") %>;
            var core_desktop_usuWP = <%= Request.Cookies("usu_WP") %>;
        
            class Color {
                constructor(r, g, b) {
                    this.set(r, g, b);
                }
            
                toString() {
                    return `rgb(${Math.round(this.r)}, ${Math.round(this.g)}, ${Math.round(this.b)})`;
                }

                set(r, g, b) {
                    this.r = this.clamp(r);
                    this.g = this.clamp(g);
                    this.b = this.clamp(b);
                }

                hueRotate(angle = 0) {
                    angle = angle / 180 * Math.PI;
                    const sin = Math.sin(angle);
                    const cos = Math.cos(angle);

                    this.multiply([
                        0.213 + cos * 0.787 - sin * 0.213,
                        0.715 - cos * 0.715 - sin * 0.715,
                        0.072 - cos * 0.072 + sin * 0.928,
                        0.213 - cos * 0.213 + sin * 0.143,
                        0.715 + cos * 0.285 + sin * 0.140,
                        0.072 - cos * 0.072 - sin * 0.283,
                        0.213 - cos * 0.213 - sin * 0.787,
                        0.715 - cos * 0.715 + sin * 0.715,
                        0.072 + cos * 0.928 + sin * 0.072,
                    ]);
                }

                grayscale(value = 1) {
                    this.multiply([
                        0.2126 + 0.7874 * (1 - value),
                        0.7152 - 0.7152 * (1 - value),
                        0.0722 - 0.0722 * (1 - value),
                        0.2126 - 0.2126 * (1 - value),
                        0.7152 + 0.2848 * (1 - value),
                        0.0722 - 0.0722 * (1 - value),
                        0.2126 - 0.2126 * (1 - value),
                        0.7152 - 0.7152 * (1 - value),
                        0.0722 + 0.9278 * (1 - value),
                    ]);
                }

                sepia(value = 1) {
                    this.multiply([
                        0.393 + 0.607 * (1 - value),
                        0.769 - 0.769 * (1 - value),
                        0.189 - 0.189 * (1 - value),
                        0.349 - 0.349 * (1 - value),
                        0.686 + 0.314 * (1 - value),
                        0.168 - 0.168 * (1 - value),
                        0.272 - 0.272 * (1 - value),
                        0.534 - 0.534 * (1 - value),
                        0.131 + 0.869 * (1 - value),
                    ]);
                }

                saturate(value = 1) {
                    this.multiply([
                        0.213 + 0.787 * value,
                        0.715 - 0.715 * value,
                        0.072 - 0.072 * value,
                        0.213 - 0.213 * value,
                        0.715 + 0.285 * value,
                        0.072 - 0.072 * value,
                        0.213 - 0.213 * value,
                        0.715 - 0.715 * value,
                        0.072 + 0.928 * value,
                    ]);
                }

                multiply(matrix) {
                    const newR = this.clamp(this.r * matrix[0] + this.g * matrix[1] + this.b * matrix[2]);
                    const newG = this.clamp(this.r * matrix[3] + this.g * matrix[4] + this.b * matrix[5]);
                    const newB = this.clamp(this.r * matrix[6] + this.g * matrix[7] + this.b * matrix[8]);

                    this.r = newR;
                    this.g = newG;
                    this.b = newB;
                }

                brightness(value = 1) {
                    this.linear(value);
                }

                contrast(value = 1) {
                    this.linear(value, -(0.5 * value) + 0.5);
                }

                linear(slope = 1, intercept = 0) {
                    this.r = this.clamp(this.r * slope + intercept * 255);
                    this.g = this.clamp(this.g * slope + intercept * 255);
                    this.b = this.clamp(this.b * slope + intercept * 255);
                }

                invert(value = 1) {
                    this.r = this.clamp((value + this.r / 255 * (1 - 2 * value)) * 255);
                    this.g = this.clamp((value + this.g / 255 * (1 - 2 * value)) * 255);
                    this.b = this.clamp((value + this.b / 255 * (1 - 2 * value)) * 255);
                }

                hsl() {
                    // Este código proviene de https://stackoverflow.com/a/9493060/2688027, bajo licencia CC BY-SA.
                    const r = this.r / 255;
                    const g = this.g / 255;
                    const b = this.b / 255;
                    const max = Math.max(r, g, b);
                    const min = Math.min(r, g, b);
                    let h, s, l = (max + min) / 2;

                    if (max === min) {
                        h = s = 0;
                    } else {
                        const d = max - min;
                        s = l > 0.5 ? d / (2 - max - min) : d / (max + min);

                        switch (max) {
                            case r:
                                h = (g - b) / d + (g < b ? 6 : 0);
                                break;

                            case g:
                                h = (b - r) / d + 2;
                                break;

                            case b:
                                h = (r - g) / d + 4;
                                break;
                        }

                        h /= 6;
                    }

                    return {
                        h: h * 100,
                        s: s * 100,
                        l: l * 100,
                    };
                }

                clamp(value) {
                    if (value > 255) {
                        value = 255;
                    } else if (value < 0) {
                        value = 0;
                    }

                    return value;
                }
            }

            class Solver {
                constructor(target, baseColor) {
                    this.target = target;
                    this.targetHSL = target.hsl();
                    this.reusedColor = new Color(0, 0, 0);
                }

                solve() {
                    const result = this.solveNarrow(this.solveWide());               
                    return {
                        values: result.values,
                        loss: result.loss,
                        filter: this.css(result.values),
                    };
                }

                solveWide() {
                    const A = 5;
                    const c = 15;
                    const a = [60, 180, 18000, 600, 1.2, 1.2];

                    let best = { loss: Infinity };

                    for (let i = 0; best.loss > 25 && i < 3; i++) {
                        const initial = [50, 20, 3750, 50, 100, 100];
                        const result = this.spsa(A, a, c, initial, 1000);

                        if (result.loss < best.loss) {
                            best = result;
                        }
                    }
                    return best;
                }

                solveNarrow(wide) {
                    const A = wide.loss;
                    const c = 2;
                    const A1 = A + 1;
                    const a = [0.25 * A1, 0.25 * A1, A1, 0.25 * A1, 0.2 * A1, 0.2 * A1];

                    return this.spsa(A, a, c, wide.values, 500);
                }

                spsa(A, a, c, values, iters) {
                    const alpha = 1;
                    const gamma = 0.16666666666666666;

                    let best = null;
                    let bestLoss = Infinity;
                    const deltas = new Array(6);
                    const highArgs = new Array(6);
                    const lowArgs = new Array(6);

                    for (let k = 0; k < iters; k++) {
                        const ck = c / Math.pow(k + 1, gamma);

                        for (let i = 0; i < 6; i++) {
                            deltas[i] = Math.random() > 0.5 ? 1 : -1;
                            highArgs[i] = values[i] + ck * deltas[i];
                            lowArgs[i] = values[i] - ck * deltas[i];
                        }

                        const lossDiff = this.loss(highArgs) - this.loss(lowArgs);

                        for (let i = 0; i < 6; i++) {
                            const g = lossDiff / (2 * ck) * deltas[i];
                            const ak = a[i] / Math.pow(A + k + 1, alpha);
                            values[i] = fix(values[i] - ak * g, i);
                        }

                        const loss = this.loss(values);

                        if (loss < bestLoss) {
                            best = values.slice(0);
                            bestLoss = loss;
                        }
                    }
                    
                    return { values: best, loss: bestLoss };

                    function fix(value, idx) {
                        let max = 100;

                        if (idx === 2 /* saturate */) {
                            max = 7500;
                        } else if (idx === 4 /* brightness */ || idx === 5 /* contrast */) {
                            max = 200;
                        }

                        if (idx === 3 /* hue-rotate */) {
                            if (value > max) {
                                value %= max;
                            } else if (value < 0) {
                                value = max + value % max;
                            }
                        } else if (value < 0) {
                            value = 0;
                        } else if (value > max) {
                            value = max;
                        }
                        
                        return value;
                    }
                }

                loss(filters) {
                    // Argument is array of percentages.
                    const color = this.reusedColor;
                    color.set(0, 0, 0);

                    color.invert(filters[0] / 100);
                    color.sepia(filters[1] / 100);
                    color.saturate(filters[2] / 100);
                    color.hueRotate(filters[3] * 3.6);
                    color.brightness(filters[4] / 100);
                    color.contrast(filters[5] / 100);

                    const colorHSL = color.hsl();

                    return (
                        Math.abs(color.r - this.target.r) +
                        Math.abs(color.g - this.target.g) +
                        Math.abs(color.b - this.target.b) +
                        Math.abs(colorHSL.h - this.targetHSL.h) +
                        Math.abs(colorHSL.s - this.targetHSL.s) +
                        Math.abs(colorHSL.l - this.targetHSL.l)
                    );
                }

                css(filters) {
                    function fmt(idx, multiplier = 1) {
                        return Math.round(filters[idx] * multiplier);
                    }

                    return `filter: invert(${fmt(0)}%) sepia(${fmt(1)}%) saturate(${fmt(2)}%) hue-rotate(${fmt(3, 3.6)}deg) brightness(${fmt(4)}%) contrast(${fmt(5)}%);`;
                }
            }

            function hexToRgb(hex) {
                // Expand shorthand form (e.g. "03F") to full form (e.g. "0033FF")

                const shorthandRegex = /^#?([a-f\d])([a-f\d])([a-f\d])$/i;
                hex = hex.replace(shorthandRegex, (m, r, g, b) => {                   
                    return r + r + g + g + b + b;
                });

                const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);

                return result
                    ? [
                        parseInt(result[1], 16),
                        parseInt(result[2], 16),
                        parseInt(result[3], 16),
                    ]
                : null;
            }

            function filter(hex) {
                const rgb = hexToRgb(hex);

                if (rgb.length !== 3) {
                    return 0;
                }

                var color = new Color(rgb[0], rgb[1], rgb[2]);
                var solver = new Solver(color);
                var result = solver.solve();

                return result.filter;
            }  
            
            function reemplazarEstiloPorClase(nombreClase, nuevoEstilo) {
                const elementos = document.querySelectorAll(`.${nombreClase}`);

                elementos.forEach(elemento => {
                    elemento.setAttribute('style', nuevoEstilo);
                });

                /*
                    Ejemplo de uso

                    reemplazarEstiloPorClase('mb_icon', 'color: green; background-color: white; font-size: 16px; border: 2px solid black;');
                */
            }

            function cambiarElementoDeEstiloPorClase(nombreClase, nuevoElementoDeEstilo) {
                const elementos = document.querySelectorAll(`.${nombreClase}`);

                elementos.forEach(elemento => {
                    for (const propiedad in nuevoElementoDeEstilo) {
                        elemento.style[propiedad] = nuevoElementoDeEstilo[propiedad];
                    }
                });

                /* 
                    Ejemplo de uso

                    cambiarElementoDeEstiloPorClase('nombre_clase', {
                        color: 'blue',                      // Cambiar color de texto
                        backgroundColor: 'lightgray',       // Cambiar color de fondo
                        fontSize: '18px'                    // Cambiar tamaño de fuente
                    });            
                */
            }  
            
            function sumarUnDia(fechaStr) {
                /*
                    IMPORTANTE:

                    Esta funcion, tal como aparece a continuación, 
                    sólo funciona con el formato de fecha "dd/MM/aaaa"

                    Si el formato de fecha en la aplicación cambia,
                    la función fallará en su respuesta.
                */

                let partes = fechaStr.split("/");
                let dia = parseInt(partes[0], 10);
                let mes = parseInt(partes[1], 10) - 1; // Restamos 1 porque los meses en JavaScript van de 0 a 11
                let anio = parseInt(partes[2], 10);

                // Crear objeto Date y sumar un día
                let fecha = new Date(anio, mes, dia);
                fecha.setDate(fecha.getDate() + 1);

                // Formatear la nueva fecha en "dd/MM/aaaa"
                let nuevoDia = fecha.getDate().toString().padStart(2, "0");
                let nuevoMes = (fecha.getMonth() + 1).toString().padStart(2, "0");
                let nuevoAnio = fecha.getFullYear();

                return `${nuevoDia}/${nuevoMes}/${nuevoAnio}`;
            }

            function restarUnDia(fechaStr) {
                /*
                    IMPORTANTE:

                    Esta funcion, tal como aparece a continuación, 
                    sólo funciona con el formato de fecha "dd/MM/aaaa"
                    
                    Si el formato de fecha en la aplicación cambia,
                    la función fallará en su respuesta.
                */

                let partes = fechaStr.split("/");
                let dia = parseInt(partes[0], 10);
                let mes = parseInt(partes[1], 10) - 1; // Restamos 1 porque los meses en JavaScript van de 0 a 11
                let anio = parseInt(partes[2], 10);

                // Crear objeto Date y restar un día
                let fecha = new Date(anio, mes, dia);
                fecha.setDate(fecha.getDate() - 1);

                // Formatear la nueva fecha en "dd/MM/aaaa"
                let nuevoDia = fecha.getDate().toString().padStart(2, "0");
                let nuevoMes = (fecha.getMonth() + 1).toString().padStart(2, "0");
                let nuevoAnio = fecha.getFullYear();

                return `${nuevoDia}/${nuevoMes}/${nuevoAnio}`;
            }         
            
            function sumar30Minutos(hora) {
                /*
                    IMPORTANTE:

                    Esta funcion, tal como aparece a continuación, 
                    sólo funciona con el formato de hora es "HH:mm"
                    (o sea, formato de 24 horas con "0" como padding)
                    
                    Si el formato de hora en la aplicación cambia,
                    la función fallará en su respuesta.

                    La hora siempre se redondea a los 30 minutos
                */

                let [hh, mm] = hora.split(":").map(Number);
                let fecha = new Date();

                if (mm <= 15) { 
                    mm = 0; 
                } else {
                    mm = 30;
                };

                fecha.setHours(hh, mm);
                fecha.setMinutes(fecha.getMinutes() + 30);
                
                return fecha.toTimeString().slice(0, 5);
            }

            function restar30Minutos(hora) {
                /*
                    IMPORTANTE:

                    Esta funcion, tal como aparece a continuación, 
                    sólo funciona con el formato de hora es "HH:mm"
                    (o sea, formato de 24 horas con "0" como padding)
                    
                    Si el formato de hora en la aplicación cambia,
                    la función fallará en su respuesta.

                    La hora siempre se redondea a los 30 minutos
                */
                               
                let [hh, mm] = hora.split(":").map(Number);
                let fecha = new Date();
                
                if (mm <= 15) { 
                    mm = 0; 
                } else {
                    mm = 30;
                };

                fecha.setHours(hh, mm);
                fecha.setMinutes(fecha.getMinutes() - 30);
                
                return fecha.toTimeString().slice(0, 5);
            }        
            
            function mask(input, mascaras) {
                /*
                    Los marcadores son:

                    9       Numérico Obligatorio
                    0       Numérico Opcional
                    A       Alfabético Obligatorio
                    _       Cualquier caracter (validar tamaño)
                    /-:.,   Literales válidos obligatorios
                    Esp     Espacio en blanco obligatorio
                */

                input.addEventListener('input', function () {

                    // Si alguna máscara usa "_", NO debemos filtrar caracteres.
                    const usaGuionBajo = mascaras.some(m => m.includes('_'));

                    // Si se usa "_", trabajamos con el valor original completo.
                    // Si no, mantenemos la regla previa (aceptar sólo alfanuméricos).
                    const raw = usaGuionBajo
                        ? input.value
                        : input.value.replace(/[^a-zA-Z0-9]/g, '');

                    const literalesPermitidos = ' /-:.,';
                    let mejorResultado = '';
                    let mayorLongitud = 0;

                    for (let mascara of mascaras) {
                        let resultado = '';
                        let i = 0; // índice del raw
                        let j = 0; // índice de la máscara

                        while (j < mascara.length) {
                            const m = mascara[j];
                            const c = raw[i];

                            if (m === '9') {
                                if (c && /\d/.test(c)) {
                                    resultado += c;
                                    i++; j++;
                                } else break;

                            } else if (m === '0') {
                                if (c && /\d/.test(c)) {
                                    resultado += c;
                                    i++;
                                }
                                j++;

                            } else if (m === 'A') {
                                if (c && /[a-zA-Z0-9]/.test(c)) {
                                    resultado += c.toUpperCase();
                                    i++; j++;
                                } else break;

                            } else if (m === '_') {
                                if (c !== undefined) {
                                    resultado += c;
                                    i++;
                                }
                                j++;

                            } else if (literalesPermitidos.includes(m)) {
                                resultado += m;
                                j++;

                            } else {
                                resultado = '';
                                break;
                            }
                        }

                        if (resultado.length > mayorLongitud) {
                            mejorResultado = resultado;
                            mayorLongitud = resultado.length;
                        }
                    }

                    input.value = mejorResultado;
                });
            }            

            function fecha_Parse(str) {
                const partes = str.split('/');
                if (partes.length !== 3) return null;
                const dia = parseInt(partes[0], 10);
                const mes = parseInt(partes[1], 10) - 1; // JS cuenta los meses desde 0
                const anio = parseInt(partes[2], 10);
                return new Date(anio, mes, dia);
            }

            function fecha_Format(fecha) {
                const dia = fecha.getDate().toString().padStart(2, '0');
                const mes = (fecha.getMonth() + 1).toString().padStart(2, '0');
                const anio = fecha.getFullYear();
                return `${dia}/${mes}/${anio}`;
            }

            function fecha_StrToDate(cadena) {
                const partes = cadena.split('/');
                if (partes.length !== 3) return null;

                const dia = parseInt(partes[0], 10);
                const mes = parseInt(partes[1], 10) - 1; // Los meses van de 0 a 11
                const anio = parseInt(partes[2], 10);

                if (isNaN(dia) || isNaN(mes) || isNaN(anio)) return null;

                return new Date(anio, mes, dia);
            }

            function formPack(formId) {
                /*
                    Esta función es especial...
                    Toma todos los valores de un FORM y los
                    "empaqueta" en una cadena larga con el formato:

                    "[Campo=valor][Campo=valor][Campo=valor][Campo=valor][Campo=valor]"

                    De esa forma podemos grabarla en un campo (varchar(MAX)) o 
                    enviarla como QueryString a un formulario que la procese
                */

                const form = document.getElementById(formId);
                const data = new FormData(form);

                let packed = "";

                for (const [key, value] of data.entries()) {
                    packed += `[${key}=${value}]`;
                }
                return packed;
            }           
               
        </script>

        <div id="mb_MenuBar" style="display:<% 
            if (NoVerEncabezado() = 1) Or (thisNoMenu = True) then    
                response.write "none"
            else  
                response.write "block"
            end if
         %>;">

            <!--
                MENU BAR MODERNO
            -->

            <div id="fw_areaMenu" class="framework_mb sombra"
                 style="display: flex; align-items: center; justify-content: space-between; 
                        height: 60px; padding: 0 16px; 
                        background-color: <%= title_background_color %>; 
                        color: <%= logo_foreground_color %>;">                        
            
                <div id="fw_divLogoMenuBar" style="display: flex; align-items: center; gap: 12px;">
                    <div id="fw_areaLogo" class="framework_logo">
                        <img src="/core/imagenes/logo_fab.png" height="46" class="title_logo" style="filter: none;">
                    </div>

                    <nav class="framework_menu" style="background-color: <%= framework_mb_background() %>;">
                        <ul id="fw_areaMenuBar" class="nav-menu" style=" <%= styleMenuBar %>">
                            <%
                                DibujarMenuBar Request.Cookies("usuario")
                            %>
                        </ul>
                    </nav>
                </div>

                <div id="fw_areaIconos" class="framework_shortcuts" style="display: flex; align-items: center; ">
                    <% 
                        if NombrePagina() = "desktop.asp" then 
                            IconosMB    
                    %>
                            <a onclick="desktop_anterior()">
                                <img class="mb_icon" src="/core/imagenes/antes.png"   height="35" style='filter: none;'>
                            </a>    

                            <a onclick="mb_toggle_min()">
                                <img class="mb_icon" src="/core/imagenes/more.png"    height="35" style='filter: none;'>
                            </a>       

                            <a onclick="desktop_siguiente()">
                                <img class="mb_icon" src="/core/imagenes/despues.png" height="35" style='filter: none;'>
                            </a>  

                        <% 
                            else 
                                '
                                ' Verificamos el estado del proceso...
                                '
                                IconosMB

                                if EsAnclable() = 1 then
                                    if EstaAnclado() = 1 then
                                        %>
                                            <a onclick="anclar('<%= thisSystem %>','<%= thisProcess %>', 0)">
                                                <img class="mb_icon" src="/core/imagenes/remove.png" width="40px" style='filter: none;' class='mb_icon'>
                                            </a>
                                        <%
                                    Else
                                        %>
                                            <a onclick="anclar('<%= thisSystem %>','<%= thisProcess %>', 1)">
                                                <img class="mb_icon" src="/core/imagenes/add.png"  width="40px" style='filter: none;' class='mb_icon'>
                                            </a>                                        
                                        <%
                                    end if
                                end if
                            %>

                            <a onclick="mb_toggle_min()">
                                <img class="mb_icon"src="/core/imagenes/more.png" height="40px" width="40px" style='filter: none;' class='mb_icon'>
                            </a>   
                    <% end if %>                
                </div>
            </div>            

            <!-- FIN DE MENU BAR -->
        </div>

        <!--
            MENU BAR "MINI""
        -->        

        <div class="framework_mb sombra" id="mb_More" style="display:<% 
            if NoVerEncabezado() = 1 then
                response.write "block"
            else  
                response.write "none"
            end if
         %>;">
         
            <div class="framework_shortcuts"
                 style="display: flex; justify-content: flex-end;align-items: center; height: 16px; padding: 0 10px;
                        background-color: <%= title_background_color %>; color: <%= icon_foreground_color %>; gap: 10px;">
                <% if NombrePagina() = "desktop.asp" then %>
                        <a onclick="desktop_anterior()">
                            <img src="/core/imagenes/antes.png"   height="18" width= "25" style='filter: none;' class='mb_icon'>
                        </a>    

                        <a onclick="mb_toggle_max()">
                            <img src="/core/imagenes/more_mini.png"    height="18" style='filter: none;' class='mb_icon'>
                        </a>       

                        <a onclick="desktop_siguiente()">
                            <img src="/core/imagenes/despues.png" height="18" width= "25" style='filter: none;' class='mb_icon'>
                        </a>                            

                <% 
                    else 
                        if EsAnclable() = 1 then
                            if EstaAnclado() = 1 then
                %>
                                <a onclick="anclar('<%= thisSystem %>','<%= thisProcess %>', 0)">
                                    <img src="/core/imagenes/remove_mini.png" width= "18" height="10" style='filter: none;' class='mb_icon'>
                                </a>
                            <% Else %>
                                <a onclick="anclar('<%= thisSystem %>','<%= thisProcess %>', 1)">
                                    <img src="/core/imagenes/add_mini.png"  width= "18" height="10" style='filter: none;' class='mb_icon'>
                                </a>                                      
                            <% end if
                        end if %>

                        <a onclick="mb_toggle_max()">
                            <img src="/core/imagenes/more_mini.png" height="18" style='filter: none;' class='mb_icon'>
                        </a>   

                        &nbsp; 
                <% end if %>
            </div>
        </div>   

        <!--
            FIN DE MENU BAR "MINI"
        -->

        <script>
            /*
                Cambiamos el color de calquier ícono en la barra de menúes
                para que corresponda con el color seleccionado por el usuario
                para las letras del menú
            */

            reemplazarEstiloPorClase("title_icon", filter("<%= icon_foreground_color %>")); 
            reemplazarEstiloPorClase("title_logo", filter("<%= logo_foreground_color %>")); 
            reemplazarEstiloPorClase("mb_icon", filter("<%= framework_mb_icon_color() %>")); 
        </script>  

        <script>
            let wBase = <%
                            if NoVerEncabezado() = 1 then
                                response.write "80"
                            else
                                response.write "21"
                            end if
                        %>;  
            var porTabla = 60;
            var porGrafica = 40;          
            var wReserva = 125;
            var maxZindex = 10;

            function Pause(secs) {
                var milliseconds = secs * 1000;
                var start = new Date().getTime();

                for (var i = 0; i < 1e7; i++) {
                    if ((new Date().getTime() - start) > milliseconds){
                    break;
                    }
                }
            }

            function padding(cadena, tama) {
                var res, stam, dif;

                res = cadena;
                stam = cadena.length;
                
                if (stam < tama) {
                    dif = tama - stam;

                    for (var i = 0; i < dif; i++) {
                        res += "\u00A0";
                    }
                }

                return res.substring(0, tama);
            }           

            function salir() {
                var confirmacion = confirm("¿Quiere salir de la aplicación?");

                if (confirmacion) {
                    window.location.href = "/forma/desktop/logout.asp";
                }
            }

            function escritorio() {
                window.location.href = "/core/desktop.asp";        
            }      

            function buscar() {
                window.location.href = "/core/buscar.asp";        
            }   

            function cola() {
                window.location.href = "/core/cola/lista.asp";           
            }

            function anclar(Sistema, Proceso, sw) {
                var vinculo = "/core/anclar.asp?s=" + Sistema + "&p=" + Proceso + "&a=" + sw;

                if (sw == 0) {
                    var confirmacion = confirm("¿Quiere eliminar este vinculo?");

                    if (confirmacion) {
                        window.location.href = vinculo;
                    }                  
                } else {
                    var confirmacion = confirm("¿Quiere agregar este vinculo?");
                    if (confirmacion) {
                        window.location.href = vinculo;
                    }                  
                }
            }

            function abrirShortcut(shortcut){
                window.location.href = shortcut;
            }      

            function profile() {
                window.location.href = "/core/profile.asp";                     
            }

            function home(){
                window.location.href = "<%= HomePage() %>";       
            }    
            
            function mb_toggle_min() {
                var mb1 = document.getElementById("mb_MenuBar");
                var mb2 = document.getElementById("mb_More");

                mb1.style.display="none";
                mb2.style.display = "block";
                menuReserva = 21;

                mb_Minimizar();
            }

            function mb_toggle_max() {
                var mb1 = document.getElementById("mb_MenuBar");
                var mb2 = document.getElementById("mb_More");

                mb1.style.display="block";
                mb2.style.display = "none";  
                menuReserva = 80;

                mb_Maximizar();
            }

            function desktop_anterior(){
                var vinculo = "/core/includes/desktop_wp.asp?wp=";
                var imagen, video;

                core_desktop_usuWP--;

                if (core_desktop_usuWP <= 0) {
                    core_desktop_usuWP = core_desktop_maxWP;
                }

                imagen = "/core/imagenes/fondos/" + core_desktop_usuWP.toString().padStart(8, '0') + ".jpg";
                video = "/core/imagenes/fondos/" + core_desktop_usuWP.toString().padStart(8, '0') + ".mp4";

                vinculo += core_desktop_usuWP;
                navigator.sendBeacon(vinculo);   

                for (let hoja of document.styleSheets) {
                    for (let regla of hoja.cssRules) {
                        if (regla.selectorText === '.bg') {
                            regla.style.backgroundImage = `url("${imagen}")`;
                        break;
                        }
                    }
                } 

                cambiarVideo(video);                
                Pause(2);
            }    

            function desktop_siguiente(){
                var vinculo ="/core/includes/desktop_wp.asp?wp=";
                var imagen, video;

                core_desktop_usuWP++;

                if (core_desktop_usuWP > core_desktop_maxWP) {
                    core_desktop_usuWP = 1;
                }                

                imagen = "/core/imagenes/fondos/" + core_desktop_usuWP.toString().padStart(8, '0') + ".jpg";
                video = "/core/imagenes/fondos/" + core_desktop_usuWP.toString().padStart(8, '0') + ".mp4";

                vinculo += core_desktop_usuWP;
                navigator.sendBeacon(vinculo);   
                
                for (let hoja of document.styleSheets) {
                    for (let regla of hoja.cssRules) {
                        if (regla.selectorText === '.bg') {
                            regla.style.backgroundImage = `url("${imagen}")`;
                            break;
                        }
                    }
                }  

                cambiarVideo(video);
                Pause(2);
            }   
            
            function cambiarVideo(ruta) {
                const video = document.getElementById('myVideo');
                const source = video.querySelector('source');

                video.pause(); // Por si estaba corriendo
                source.src = ruta;
                video.load();
                video.play().catch(() => {
                    console.warn("No se pudo reproducir automáticamente.");
                });
            }            
            
            function esPortable() {
                // Método de la Cadena del Agente

                let isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);

                // Método de la Resolución de la Pantalla

                if (!isMobile) {
                    let screenWidth = window.screen.width;
                    let screenHeight = window.screen.height;
                    isMobile = (screenWidth < 768 || screenHeight < 768);
                }

                // Método de los Eventos Táctiles
                
                if (!isMobile) {
                    isMobile = (('ontouchstart' in window) || (navigator.maxTouchPoints > 0) || (navigator.msMaxTouchPoints > 0));
                }

                // Método de Consultas de Medios

                if (!isMobile) {
                    let bodyElement = document.getElementsByTagName('body')[0];
                    isMobile = window.getComputedStyle(bodyElement).getPropertyValue('content').indexOf('mobile') !== -1;
                }

                return isMobile
            }        

            function esTelefono() {
                // Método de la Cadena del Agente

                let isMobile = /Android|iPhone/i.test(navigator.userAgent);

                // Método de la Resolución de la Pantalla

                if (!isMobile) {
                    let screenWidth = window.screen.width;
                    let screenHeight = window.screen.height;
                    isMobile = (screenWidth < 768 || screenHeight < 768);
                }


                return isMobile
            }           
        </script>

        <!-- #include virtual = "/forma/plus/plusPerfiles.plus" -->
        <% mb_conn.close: set mb_conn = nothing %>        