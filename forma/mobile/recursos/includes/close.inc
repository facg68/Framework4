<script>
    // Motor de Sonidos para Notificaciones ----------
        const sonidos = {
            A: new Audio('/core/sounds/alert_00004.mp3'),
            E: new Audio('/core/sounds/alert_00005.mp3'),
            P: new Audio('/core/sounds/alert_00010.mp3'),
        };

        (function habilitarSonidosMovil() {
            document.addEventListener('touchstart', function () {

                const dummy = new Audio('/core/sounds/alert_00000.mp3');
                dummy.play().catch(() => {});

            }, { once: true });
        })();    
    //------------------------------------------------

    // Motor de Notificaciones -----------------------
        function checkNotificaciones() {
            fetch('/core/includes/ajax/notificaciones.asp')
                .then(res => res.json())
                .then(data => {
                    if (Array.isArray(data) && data.length > 0) {
                        mostrarNotificaciones(data);
                    }
                })
                .catch(() => {
                    /* Silencio móvil: no molestamos al usuario */
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

                    /* En móvil va centrado, no toast */
                    position: 'center',
                    toast: false,
                    icon: 'info',
                    showConfirmButton: false,
                    timer: 4500,
                    timerProgressBar: true,
                    background: '#f2f2f2',

                    /* Que se vea como F4 */
                    customClass: {
                        popup: 'swal-f4-popup',
                        title: 'swal-f4-title',
                        htmlContainer: 'swal-f4-text'
                    }
                };

                if (n.tipo === 'E') {
                    config.icon = 'error';
                    config.background = '#ffe6e6';
                }
                else if (n.tipo === 'P') {
                    config.icon = 'warning';
                    config.background = '#fff3cd';
                }

                sonidos[n.tipo].play().catch(() => {});
                Swal.fire(config);

                fetch('/core/includes/ajax/notificaciones_marcar.asp?s=' + n.secuencia)
                    .catch(() => {});

                setTimeout(() => {
                    mostrarUna(i + 1);
                }, config.timer + 300);
            };
            mostrarUna(0);
        }

        setInterval(checkNotificaciones, 90000);
    //------------------------------------------------

    // Motor de Objetos Plus (Overlay + Centrados) ---
        function cerrarObjetosPlus() {
            const overlay = document.querySelector(".plus-overlay");
            if (overlay) {
                overlay.classList.remove("active");
            }

            document.querySelectorAll('.objeto-plus').forEach(menu => {
                menu.classList.remove('plus-visible');
            });
        }

        function mostrarObjetoPlus(trigger, idMenu) {
            cerrarObjetosPlus();

            const menu = document.getElementById(idMenu);
            if (!menu) return;

            menu.dataset.origen = trigger.id || '';

            if (!ejecutarPlusInit(idMenu)) {
                return;
            }

            const overlay = document.querySelector(".plus-overlay");
            if (overlay) {
                overlay.classList.add("active");
            }

            menu.classList.add('plus-visible');
        }

        function ejecutarPlusInit(idMenu) {
            const nombreFuncion = idMenu + "_Init";
            const fn = window[nombreFuncion];

            if (typeof fn === "function") {
                return fn(idMenu);
            }

            return true;
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

                /* ============================= */
                /* 1️⃣ Separar modo visual       */
                /* ============================= */

                const partesModo = clean.split('|');
                const definicion = partesModo[0].trim();
                const modoRaw = partesModo[1] ? partesModo[1].trim() : "lista";
                const modo = (modoRaw === "botones") ? "botones" : "lista";

                /* ============================= */
                /* 2️⃣ Separar selector y plus  */
                /* ============================= */

                const partes = definicion.split(':');
                if (partes.length < 2) return;

                let selector = partes[0].trim();
                const plusID = partes[1].trim();

                if (!selector.startsWith('#') && !selector.startsWith('.')) {
                    selector = '#' + selector;
                }

                const menu = document.getElementById(plusID);
                if (!menu) return;

                document.querySelectorAll(selector).forEach(trigger => {
                    trigger.addEventListener("click", function(e) {
                        e.stopPropagation();

                        // Aplicar modo visual dinámicamente
                        menu.classList.remove("plus-modo-lista", "plus-modo-botones");
                        menu.classList.add("plus-modo-" + modo);

                        mostrarObjetoPlus(trigger, plusID);
                    });
                });
            });
        });

        /* Cierre global */

        document.addEventListener('click', function(e) {
            // Click en opción
            if (e.target.closest('.objeto-plus .plus-item')) {
                cerrarObjetosPlus();
                return;
            }

            // Click fuera
            if (!e.target.closest('.objeto-plus')) {
                cerrarObjetosPlus();
            }
        });

        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape') {
                cerrarObjetosPlus();
            }
        });
    //------------------------------------------------

    // Motor de Menú Lateral -------------------------
        function toggleMenu() {
            document.getElementById('sidebar').classList.toggle('open');
            document.body.classList.toggle('menu-open');
        }    
    //------------------------------------------------

    // Guardar Cookies -------------------------------
        function guardarCookie(nombre, valor) {
            var fecha = new Date();
            var dias = 30;

            fecha.setTime(fecha.getTime() + (dias * 24 * 60 * 60 * 1000));

            document.cookie = nombre + "=" + encodeURIComponent(valor) +
                ";expires=" + fecha.toUTCString() +
                ";path=/;SameSite=Lax";
        }    

        function borrarCookie(nombre) {
            document.cookie = nombre +
                "=;expires=Thu, 01 Jan 1970 00:00:00 UTC;path=/;SameSite=Lax";
        }   
        
        function leerCookie(nombre) {
            const nombreEQ = nombre + "=";
            const cookies = document.cookie.split(';');

            for (let i = 0; i < cookies.length; i++) {
                let c = cookies[i].trim();
                if (c.indexOf(nombreEQ) === 0) {
                    return c.substring(nombreEQ.length);
                }
            }
            return null;
        }              
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
</script>

<div id="rotar-mensaje">
    Rote el periférico 90° para continuar
</div>