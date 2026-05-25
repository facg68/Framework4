        <script>
            let anchoLogoBarOriginal = 0;
            let anchoIconosOriginal = 0;
            let sumaNecesaria = 0;

            const sonidos = {
                A: new Audio('/core/sounds/alert_00004.mp3'),
                E: new Audio('/core/sounds/alert_00005.mp3'),
                P: new Audio('/core/sounds/alert_00010.mp3')
            };

            // Lectura de Parámetros desde <body> ------------
                function leerAtributoPorcentaje(nombreAttr, defecto) {
                    if (document.body.hasAttribute(nombreAttr)) {
                        let v = parseInt(document.body.getAttribute(nombreAttr), 10);

                        if (!isNaN(v)) {
                            return Math.max(10, Math.min(v, 100));
                        }
                    }
                    return defecto;
                }                
                
                function leerAtributoEntero(nombreAttr, defecto) {
                    if (document.body.hasAttribute(nombreAttr)) {
                        let v = parseInt(document.body.getAttribute(nombreAttr), 10);

                        if (!isNaN(v) && v >= 0 && v <= 500) {
                            return v;
                        }
                    }
                    return defecto;
                }                

                let plantilla = (document.body.getAttribute("plantilla") || "Normal").toLowerCase();
                pageReserva = leerAtributoEntero("reserva", 200);      

                // Flags

                let tieneTabla      = document.body.hasAttribute("tabla");
                let tieneGrafica    = document.body.hasAttribute("grafica");

                // Leer porcentajes personalizados (si aplica)

                let tablaPorcentaje   = leerAtributoPorcentaje("tabla", 100);
                let graficaPorcentaje = leerAtributoPorcentaje("grafica", 100);

                // Si solo definió uno, calculamos el otro

                if (tieneTabla && !tieneGrafica) {
                    graficaPorcentaje = 100 - tablaPorcentaje;
                }

                if (!tieneTabla && tieneGrafica) {
                    tablaPorcentaje = 100 - graficaPorcentaje;
                }

                // Si no se especificaron proporciones o la suma supera el 100%, aplicar valores por defecto

                if (plantilla == "dividida") {
                    const algunoEsCero = (tablaPorcentaje === 0 || graficaPorcentaje === 0);
                    const sumaInvalida = (tablaPorcentaje + graficaPorcentaje > 100);
                    const ningunoDefinido = (!tieneTabla && !tieneGrafica);

                    if (ningunoDefinido || sumaInvalida || algunoEsCero) {
                        tablaPorcentaje = 55;
                        graficaPorcentaje = 45;
                    }
                }                
            //------------------------------------------------  

            // Lógica de Inicialización de Objetos------------
                switch (plantilla) {
                    case "lista":        InicializarLista(); break;
                    case "tabla":        InicializarTabla(tablaPorcentaje); break;
                    case "grafica":      InicializarGrafica(graficaPorcentaje); break;
                    case "dividida":     InicializarTablaGrafica(tablaPorcentaje, graficaPorcentaje); break;
                    case "panoramica":   InicializarPanoramicChart(graficaPorcentaje); break;
                    case "area":         InicializarAreaChart(graficaPorcentaje); break;
                    case "desktop":      InicializarDesktop(); break;
                    default:             InicializarPagina(); break;
                }            
            //------------------------------------------------

            // Helpers para todos los casos ------------------
                function calcularPixelesDesdePorcentaje(porcentaje) {
                    if (porcentaje <= 0 || porcentaje > 100) {
                        console.error('El porcentaje debe estar entre 1 y 100');
                        return 0;
                    }

                    const altoVentana = window.innerHeight - pageReserva - menuReserva;
                    return (altoVentana * porcentaje) / 100;
                }      
                
                function ajustarGraficaPorPixeles(newHeight) {
                    chart.updateOptions({
                        chart: { height: newHeight }
                    });
                }    

                function ajustarTablaPorPixeles(pixels) {
                    const contenedor = document.querySelector('.tabla-wrapper');
                    if (!contenedor) return;

                    contenedor.style.height = pixels + 'px';
                    contenedor.style.overflowY = 'auto';
                } 

                function existeClase(compuesta) {
                    return document.querySelector('.' + compuesta) !== null;
                }  

                function medirDimensionesIniciales() {
                    const logoMenuBar = document.getElementById("fw_divLogoMenuBar");
                    const areaIconos = document.getElementById("fw_areaIconos");

                    if (!logoMenuBar || !areaIconos) return;

                    // Medimos y guardamos anchos REALES al cargar

                    anchoLogoBarOriginal = logoMenuBar.offsetWidth + 45;
                    anchoIconosOriginal = areaIconos.offsetWidth + 45;
                    sumaNecesaria = anchoLogoBarOriginal + anchoIconosOriginal;
                }

                function ajustarAreaMenu() {
                    const menuCompleto = document.getElementById("mb_MenuBar");
                    const arealogo = document.getElementById("fw_areaLogo");
                    const areaIconos = document.getElementById("fw_areaIconos");

                    if (!menuCompleto || !arealogo || !areaIconos) return;

                    // Siempre los mostramos antes de evaluar

                    arealogo.style.display = "";
                    areaIconos.style.display = "";

                    requestAnimationFrame(() => {
                        const anchoTotal = menuCompleto.offsetWidth;

                        if (anchoTotal < sumaNecesaria) {
                            areaIconos.style.display = "none";
                        }

                        if (anchoTotal < anchoLogoBarOriginal) {
                            arealogo.style.display = "none";
                        }
                    });
                }

                window.addEventListener("load", () => {
                    medirDimensionesIniciales();
                    ajustarAreaMenu();
                });

                window.addEventListener("resize", ajustarAreaMenu);    
                
                // Verficar que no se abran Plus cuando estamos en el menú

                const areaMenu = document.getElementById("fw_areaMenuBar");

                areaMenu.addEventListener("mouseenter", function() {
                    cerrarObjetosPlus();
                });              
            //------------------------------------------------  

            // Normal (main, main-scoll) ----------------------
                function InicializarPagina() {
                    window.addEventListener("resize", ajustarMain);
                    ajustarMain();

                    mb_Maximizar = () => ajustarMain();
                    mb_Minimizar = () => ajustarMain();
                }  

                function ajustarMain() {
                    var espDisponible = window.innerHeight - pageReserva - menuReserva;
                    const scrollMain = document.querySelector('.main-scroll');

                    if (scrollMain) {
                        scrollMain.style.maxHeight = espDisponible + 'px';
                    }
                }                   
            //------------------------------------------------

            // Listas (overflow) -----------------------------
                function InicializarLista() {
                    window.addEventListener("resize", ajustarLista);
                    ajustarLista();
                    
                    mb_Maximizar = () => ajustarLista();
                    mb_Minimizar = () => ajustarLista();
                }  

                function ajustarLista() {
                    var espDisponible = window.innerHeight - pageReserva - menuReserva;

                    const miDiv = document.getElementById("overFlow");
                    if (miDiv) {
                        miDiv.style.width = "100%";
                        miDiv.style.height = espDisponible + "px";
                        miDiv.style.overflow = "auto";
                        miDiv.style.backgroundColor = "rgb(207, 207, 207)";
                    }
                }              
            //------------------------------------------------                

            // Simple: Gráfica -------------------------------
                function ajustarGfx() {
                    const pixGrafica = calcularPixelesDesdePorcentaje(porGrafica);
                    document.body.offsetHeight;
                    ajustarGraficaPorPixeles(pixGrafica);
                    document.body.offsetHeight;
                }               

                function InicializarGrafica(porcentajeGrafica) {
                    porGrafica = porcentajeGrafica;
                    window.addEventListener('resize', ajustarGfx);
                    ajustarGfx();

                    mb_Maximizar = () => ajustarGfx();
                    mb_Minimizar = () => ajustarGfx();
                }
            //------------------------------------------------      
            
            // Simple: Tabla ---------------------------------
                function ajustarTabla() {
                    const pixTabla = calcularPixelesDesdePorcentaje(porTabla);
                    document.body.offsetHeight;
                    ajustarTablaPorPixeles(pixTabla);
                    document.body.offsetHeight;
                }               

                function InicializarTabla(porcentajeTabla) {
                    porTabla = porcentajeTabla;
                    window.addEventListener('resize', ajustarTabla);
                    ajustarTabla();

                    mb_Maximizar = () => ajustarTabla();
                    mb_Minimizar = () => ajustarTabla();
                }
            //------------------------------------------------                

            // Doble: Lista + Gráfica ------------------------
                function dbl_ajustar() {
                    const pixTabla = calcularPixelesDesdePorcentaje(porTabla);
                    const pixGrafica = calcularPixelesDesdePorcentaje(porGrafica);

                    document.body.offsetHeight;
                    ajustarTablaPorPixeles(pixTabla);
                    ajustarGraficaPorPixeles(pixGrafica);
                    document.body.offsetHeight;
                }

                function InicializarTablaGrafica(porcentajeTabla, porcentajeGrafica) {
                    porTabla = porcentajeTabla;
                    porGrafica = porcentajeGrafica;

                    window.addEventListener('resize', dbl_ajustar);
                    dbl_ajustar();

                    mb_Maximizar = () => dbl_ajustar();
                    mb_Minimizar = () => dbl_ajustar();
                }
            //------------------------------------------------        
            
            // Panoramic Chart: 2 Gráficas -------------------
                function pan_ajustar() {
                    var espDisponible = window.innerHeight - pageReserva - menuReserva;
                    let disponible = Math.floor(espDisponible * (porGrafica / 100));

                    let bigH = Math.floor(disponible * 0.70);   // 70%
                    let panH = disponible - bigH;               // 30%

                    chart.updateOptions({
                        chart: { height: bigH }                 // Ajusta el gráfico grande
                    }, false, true);

                    chartLine.updateOptions({
                        chart: { height: panH }                 // Ajusta el gráfico panorámico
                    }, false, true);
                }

                function InicializarPanoramicChart(porcentajeGrafica) {
                    porGrafica = porcentajeGrafica;
                    window.addEventListener('resize', pan_ajustar);
                    pan_ajustar();

                    // Ajustar comportamiento de botones
                    mb_Maximizar = () => pan_ajustar();
                    mb_Minimizar = () => pan_ajustar();
                }
            //------------------------------------------------      
            
            // Area Chart: -----------------------------------
                function area_Ajustar() {
                    let winH = window.innerHeight - pageReserva - menuReserva;
                    let disponible = Math.floor(winH * (porGrafica / 100));

                    chart.updateOptions({
                        chart: { height: disponible }       // Ajusta el gráfico
                    }, false, true);
                }

                function InicializarAreaChart(porcentajeGrafica) {
                    porGrafica = porcentajeGrafica;
                    window.addEventListener('resize', area_Ajustar);
                    area_Ajustar();

                    mb_Maximizar = () => area_Ajustar();
                    mb_Minimizar = () => area_Ajustar();
                }
            //------------------------------------------------   
            
            // Desktop: --------------------------------------
                function InicializarDesktop() {
                    /* 
                        Reservada para más adelante 
                    */
                }
            //------------------------------------------------
            
            // Motor de Notificaciones: ----------------------
                (function habilitarSonidosF4() {
                    document.addEventListener('click', function activarSonido() {
                        const sonda = new Audio('/core/sounds/alert_00004.mp3');
                        sonda.volume = 0;
                        sonda.play().catch(() => {});
                        document.removeEventListener('click', activarSonido);
                    }, { once: true });
                })();

                function checkNotificaciones() {
                    fetch('/core/includes/ajax/notificaciones.asp')
                        .then(res => res.json())
                        .then(data => {
                            if (data && data.length > 0) {
                                mostrarNotificaciones(data);
                            }
                        })
                        .catch(err => {
                            /* No hay notificaciones. */
                            //console.error('Error al consultar notificaciones:', err);
                        });
                }                

                function mostrarNotificaciones(lista) {
                    if (!Array.isArray(lista) || lista.length === 0) return;

                    const mostrarUna = (i) => {
                        if (i >= lista.length) return;

                        const n = lista[i];
                        let config = {
                            title: n.titulo,
                            text: n.mensaje,
                            position: 'top-end',
                            toast: true,
                            showConfirmButton: false,
                            icon: 'info',
                            timer: 4000,
                            background: '#e1eefc',
                            customClass: { title: 'swal-alerta' }
                        };

                        let sonido = '/core/sounds/alert_00004.mp3';                                // Alerta por defecto

                        if (n.tipo === 'E') {
                            config.icon = 'error';
                            config.timer = 6000;
                            config.background = '#ffe6e6';
                            config.customClass.title = 'swal-error';
                            sonido = '/core/sounds/alert_00005.mp3';
                        } else if (n.tipo === 'P') {
                            config.icon = 'warning';
                            config.timer = 8000;
                            config.background = '#fff3cd';
                            config.customClass.title = 'swal-prioritario';
                            sonido = '/core/sounds/alert_00010.mp3';
                        }

                        sonidos[n.tipo].play().catch(() => {});                                     // Reproducir sonido
                        Swal.fire(config);                                                          // Mostrar notificación visual
                        
                        if (n.vinculo && n.vinculo.trim() !== "") {                                 // Ejecutar vínculo (si aplica)
                            fetch(n.vinculo).catch(() => {});
                        }
                        
                        fetch('/core/includes/ajax/notificaciones_marcar.asp?s=' + n.secuencia);    // Marcar como entregada

                        setTimeout(() => {                                                          // Mostrar la siguiente (si hay más)
                            mostrarUna(i + 1);
                        }, config.timer + 200);
                    };

                    mostrarUna(0);
                }

                function abrirNotificaciones() {
                    const modal = document.getElementById('notificacionesModal');
                    const contenedor = document.getElementById('modalListaNotificaciones');

                    modal.classList.remove('oculto');
                    contenedor.innerHTML = '<p>Cargando notificaciones...</p>';

                    fetch('/core/includes/ajax/notificaciones_lista.asp')
                        .then(res => res.json())
                        .then(data => {
                            if (!Array.isArray(data) || data.length === 0) {
                                contenedor.innerHTML = '<p class="modal-vacio">No hay notificaciones aún.</p>';
                                return;
                            }

                            let html = '';
                            data.forEach(n => {
                                html += `
                                    <div class="modal-item tipo-${n.tipo}">
                                        <div class="item-titulo">${n.titulo}</div>
                                        <div class="item-mensaje">${n.mensaje}</div>
                                        <div class="item-fecha">#${n.secuencia}</div>
                                    </div>
                                `;
                            });

                            contenedor.innerHTML = html;
                        })
                        .catch(() => {
                            contenedor.innerHTML = '<p class="modal-error">Error al cargar notificaciones.</p>';
                        });
                }

                function cerrarNotificaciones() {
                    document.getElementById('notificacionesModal').classList.add('oculto');
                }                

                setInterval(checkNotificaciones, 60000);
            //------------------------------------------------

            // Motor de Objetos Plus: ------------------------
                function cerrarObjetosPlus() {
                    document.querySelectorAll('.objeto-plus').forEach(menu => {
                        menu.classList.remove('plus-visible');
                    });
                }

                function mostrarObjetoPlus(trigger, idMenu) {
                    cerrarObjetosPlus();

                    const menu = document.getElementById(idMenu);
                    if (!menu) return;

                    menu.dataset.origen = trigger.id || '';
                    const rect = trigger.getBoundingClientRect();

                    menu.style.top  = (rect.bottom + window.scrollY) + 'px';
                    menu.style.left = (rect.left + window.scrollX) + 'px';

                    if (ejecutarPlusInit(idMenu)) {
                        menu.classList.add('plus-visible');
                    }
                }

                function ejecutarPlusInit(idMenu) {
                    const nombreFuncion = idMenu + "_Init";
                    const fn = window[nombreFuncion];

                    if (typeof fn === "function") {
                        return fn(idMenu);
                    }

                    return true;
                }

                function registrarPlus(triggerSelector, plusID, evento, modo) {
                    const triggers = document.querySelectorAll(triggerSelector);
                    const menu = document.getElementById(plusID);

                    if (!menu) return;

                    triggers.forEach(trigger => {
                        trigger.addEventListener(evento, function(e) {
                            e.stopPropagation();

                            // Aplicar modo visual dinámicamente
                            menu.classList.remove("plus-modo-lista", "plus-modo-botones", "plus-modo-cinta");
                            menu.classList.add("plus-modo-" + modo);

                            mostrarObjetoPlus(this, plusID);
                        });
                    });
                }

                function obtenerOrigenPlus(elemento) {
                    const menu = elemento.closest('.objeto-plus');
                    return menu?.dataset?.origen || null;
                }

                document.addEventListener('DOMContentLoaded', () => {
                    const reglasRaw = document.body.getAttribute('plus') || '';
                    const reglas = reglasRaw.split(';');

                    reglas.forEach(regla => {
                        const clean = regla.trim();
                        if (!clean) return;

                        // 1️⃣ Separar modo visual
                        const partesModo = clean.split('|');
                        const definicion = partesModo[0].trim();
                        
                        const modoRaw = partesModo[1] ? partesModo[1].trim() : "lista";
                        let modo = "lista";

                        if (modoRaw === "botones")    { modo = "botones"; }
                        else if (modoRaw === "cinta") { modo = "cinta"; }

                        // 2️⃣ Separar selector y resto
                        const partesSelector = definicion.split(':');
                        if (partesSelector.length < 2) return;

                        let selector = partesSelector[0].trim();
                        const resto = partesSelector[1].trim();

                        // 3️⃣ Separar plus y evento
                        const partesEvento = resto.split('@');
                        const plusID = partesEvento[0].trim();
                        const evento = partesEvento[1] ? partesEvento[1].trim() : "click";

                        // Normalización selector
                        if (!selector.startsWith('#') && !selector.startsWith('.')) {
                            selector = '#' + selector;
                        }

                        registrarPlus(selector, plusID, evento, modo);
                    });

                });

                document.addEventListener('click', function(e) {
                    if (e.target.closest('.objeto-plus .plus-item')) {
                        cerrarObjetosPlus();
                        return;
                    }

                    if (!e.target.closest('.objeto-plus')) {
                        cerrarObjetosPlus();
                    }
                }, true); // 👈 CAPTURA                

                // Cerrar con tecla ESC
                document.addEventListener('keydown', function(e) {
                    if (e.key === 'Escape') {
                        cerrarObjetosPlus();
                    }
                });  
            //------------------------------------------------     
            
            // Motor Almacenamiento: -------------------------
                const F4Storage = (() => {

                    const prefix = "F4::";

                    const disponible = (() => {
                        try {
                            const test = "__f4test__";
                            localStorage.setItem(test, test);
                            localStorage.removeItem(test);
                            return true;
                        } catch {
                            return false;
                        }
                    })();

                    function normalize(texto) {
                        return typeof texto === "string"
                            ? texto.trim().toLowerCase()
                            : null;
                    }

                    function buildKey(usuario, llave) {
                        const u = normalize(usuario);
                        const k = normalize(llave);

                        if (!u || !k) return null;

                        return `${prefix}${u}::${k}`;
                    }

                    function isValidValue(valor) {
                        return (
                            typeof valor === "string" ||
                            typeof valor === "number" ||
                            typeof valor === "boolean"
                        );
                    }

                    return {
                        set(usuario, llave, valor) {
                            if (!disponible) return false;

                            const key = buildKey(usuario, llave);
                            if (!key || !isValidValue(valor)) return false;

                            localStorage.setItem(key, valor);
                            return true;
                        },

                        get(usuario, llave) {
                            if (!disponible) return null;

                            const key = buildKey(usuario, llave);
                            if (!key) return null;

                            return localStorage.getItem(key);
                        },

                        remove(usuario, llave) {
                            if (!disponible) return;

                            const key = buildKey(usuario, llave);
                            if (!key) return;

                            localStorage.removeItem(key);
                        },

                        exists(usuario, llave) {
                            if (!disponible) return false;

                            const key = buildKey(usuario, llave);
                            if (!key) return false;

                            return localStorage.getItem(key) !== null;
                        },

                        clearUsuario(usuario) {
                            if (!disponible) return;

                            const u = normalize(usuario);
                            if (!u) return;

                            const pattern = `${prefix}${u}::`;

                            for (let i = localStorage.length - 1; i >= 0; i--) {
                                const key = localStorage.key(i);
                                if (key && key.startsWith(pattern)) {
                                    localStorage.removeItem(key);
                                }
                            }
                        },

                        clearFramework() {
                            if (!disponible) return;

                            for (let i = localStorage.length - 1; i >= 0; i--) {
                                const key = localStorage.key(i);
                                if (key && key.startsWith(prefix)) {
                                    localStorage.removeItem(key);
                                }
                            }
                        }
                    };
                })();
            //------------------------------------------------                     

            registrarPlus('#perfil_icon', 'plus_perfiles', 'mouseenter', 'lista');            
        </script>

        <div id="notificacionesModal" class="modal-f4 oculto">
            <div class="modal-contenido">
                <div class="modal-header">
                    <span class="modal-titulo">Notificaciones</span>
                    <span class="modal-cerrar" onclick="cerrarNotificaciones()">×</span>
                </div>
                <div id="modalListaNotificaciones" class="modal-lista">
                    <p>Cargando...</p>
                </div>
            </div>
        </div>   