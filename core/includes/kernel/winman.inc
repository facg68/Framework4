        <script type="text/javascript">
            var highestZ;

            function init(PresentarSnippets) {
                highestZ = 20; // Valor inicial de z-index      

                if (PresentarSnippets === 1) {       
                    <%
                        dim Snips, SnipsSQL

                        SnipsSQL = "SELECT s.snippet, s.snippetLeft, s.snippetTop, p.proNombre, s.snippetMinimizado, " & _
                                         " '/forma/desktop/apps/' + p.proSistema + '/' + p.proAction + '.asp' AS Accion " & _
                                     "FROM dbo.seg_Usuarios_Snippets AS s " & _
                               "INNER JOIN dbo.seg_Procesos AS p " & _
                                       "ON s.codProceso = p.proCodigo " & _
                                      "AND s.codSistema = p.proSistema " & _
                                    "WHERE (s.snippetActivo = 1) " & _
                                      "AND (s.codUsuario = '" & Request.Cookies("Usuario") & "') " & _
                                 "ORDER BY s.snippetIndex;"

                        set Snips = Desktop_Conn.execute(SnipsSQL)
                            if not (Snips.bof or Snips.eof) then
                                Do
                                    %> 
                                        openWindow('<%= Snips("snippet") %>', <%= Snips("snippetTop") %>, <%= Snips("snippetLeft") %>, '<%= Snips("proNombre") %>', '<%= Snips("Accion") %>', <%= Snips("snippetMinimizado") %>); 
                                    <%
                                    Snips.MoveNext
                                Loop Until (Snips.eof)
                            end if
                        Snips.close: set Snips = nothing
                    %>
                }
            }

            // Funciones del Manejador de Ventanas ----------------------------------------------------------------------------------

                function openWindow(nombreObjeto, top, left, titulo, accion, minimizado = 0) {
                    const existing = document.getElementById("snippet_" + nombreObjeto + "_b");

                    if (existing) {
                        if (existing.dataset.minimized === "1") {
                            restoreWindow(nombreObjeto);
                        } else {
                            highestZ += 10;
                            existing.style.zIndex = highestZ;
                        }
                        return;
                    }                

                    const div = document.createElement("div");
                    highestZ += 10;

                    // Crear el Frame de la Vetnana ( body "_b" ) --------------------------------------------------------------------

                        div.id = "snippet_" + nombreObjeto + "_b";
                        div.className = "winBody"; 
                        div.style.width = "0px";
                        div.style.left = left + "px";
                        div.style.top = top + "px";
                        div.style.display = "none";
                        div.style.position = "absolute";
                        div.style.zIndex = highestZ;
                        div.dataset.titulo = titulo;

                    // Fin - Crear el Frame de la Venana -----------------------------------------------------------------------------

                    // Crear la barra de título e insertamos botones ( header "_h" ) -------------------------------------------------

                        div.innerHTML = `
                            <div class="winHeader"
                                 id="snippet_${nombreObjeto}_h"
                                 onclick="divFocus('snippet_${nombreObjeto}_b','snippet_${nombreObjeto}_h')"
                                 style="display:flex; align-items:center;
                                        background-color: var(--menu-barra-fondo);
                                        color: var(--menu-barra-font-color);">

                                <div style="flex:1; text-align:left; overflow:hidden; white-space:nowrap; text-overflow:ellipsis;
                                            font: <%= framework_menu_font_size() %>px <%= framework_mb_font() %>;
                                            color:<%= framework_mb_font_color() %>;">
                                    &nbsp;${titulo}
                                </div>

                                <div style="margin-left:auto; display:flex; gap:6px; align-items:center;">
                                    <img class="clickable window_title_button"
                                        src="/core/imagenes/min.png"
                                        onclick="event.stopPropagation(); minimizeWindow('${nombreObjeto}')"/>

                                    <img class="clickable window_title_button"
                                        src="/core/imagenes/max.png"
                                        onclick="event.stopPropagation(); maximizeWindow('${accion}')"/>

                                    <img class="clickable window_title_button"
                                        src="/core/imagenes/cerrar.png"
                                        onclick="event.stopPropagation(); closeWindow('${nombreObjeto}')"/>
                                    &nbsp;
                                </div>
                            </div>

                            <div id="snippet_${nombreObjeto}_c"></div>
                        `;

                        document.getElementById("desk_div_container").appendChild(div);
                        const header = div.querySelector(".winHeader");
                        dragWindow(header, div);

                        // Activar ventana al hacer click en cualquier parte

                        div.addEventListener("pointerdown", function(e) {
                            if (e.target.closest(".window_title_button")) return;
                            bringToFront(div);
                        });                        

                    // Fin - Crear la barra de títitulo e insertar botones -----------------------------------------------------------                        

                    navigator.sendBeacon("/core/includes/snip.asp?s=" + nombreObjeto + "&est=1");

                    // Crear el contenedor e insertar el Snippet ( content "_c" ) ----------------------------------------------------

                        fetch("/forma/snippets/" + nombreObjeto + ".asp")
                            .then(r => r.text())
                            .then(html => {

                                const body = document.getElementById("snippet_" + nombreObjeto + "_c");
                                body.innerHTML = html;

                                limpiarScriptsVentana(div);
                                registrarScriptsEnContenedor(body, div);

                                const initFunctionName = nombreObjeto + "_init";

                                let intentos = 0;

                                const tryInit = () => {
                                    if (typeof window[initFunctionName] === "function") {
                                        window[initFunctionName](div);

                                        reemplazarEstiloPorClase(
                                            "window_title_button",
                                            filter("<%= icon_foreground_color %>")
                                        );

                                        if (minimizado == 1) {
                                            minimizeWindow(nombreObjeto);
                                        } else {
                                            div.style.display = "block";
                                        }

                                    } else {
                                        intentos++;
                                        if (intentos < 25) {
                                            setTimeout(tryInit, 10);
                                        } else {
                                            console.warn("No se encontró:", initFunctionName);
                                        }
                                    }
                                };

                                tryInit();
                            })
                            .catch(err => console.error("Error cargando " + nombreObjeto, err));

                    // Fin - Crear el contetenedor e insertar el Snippet ( content "_c" ) ---------------------------------------------
                }

                function loadInWindow(nombreSnippet, url) {
                    if (!url || typeof url !== "string") {
                        console.warn("URL inválida");
                        return;
                    }

                    const body = document.getElementById("snippet_" + nombreSnippet + "_c");
                    const win  = document.getElementById("snippet_" + nombreSnippet + "_b");

                    if (!body || !win) return;
                    
                    const current = win.dataset.currentSnippet || nombreSnippet;        // Verificar el Snippet cargado
                    const closeFunctionName = current + "_close";                       // Ejecutamos el "_close()"

                    if (typeof window[closeFunctionName] === "function") {
                        try {
                            window[closeFunctionName](win);
                        } catch (e) {
                            console.warn("Error en close:", closeFunctionName, e);
                        }
                    }

                    limpiarScriptsVentana(win);                                         // Limpiar scripts asociados a la ventana

                    // Cargamos nuevo Snippet en el mismo "Frame" (ventana)

                    fetch(url)
                        .then(r => r.text())
                        .then(html => {

                            body.innerHTML = html;
                            win.dataset.currentSnippet = nombreSnippet;                 // "Registramos" nuevo snippet activo
                            registrarScriptsEnContenedor(body, win);

                            const initFunctionName = nombreSnippet + "_init";
                            let intentos = 0;

                            const tryInit = () => {
                                if (typeof window[initFunctionName] === "function") {

                                    try {
                                        window[initFunctionName](win);
                                    } catch (e) {
                                        console.error("Error en init:", initFunctionName, e);
                                    }

                                    reemplazarEstiloPorClase(
                                        "window_title_button",
                                        filter("<%= icon_foreground_color %>")
                                    );

                                } else {
                                    intentos++;
                                    if (intentos < 50) {
                                        setTimeout(tryInit, 10);
                                    } else {
                                        console.warn("No se encontró:", initFunctionName);
                                    }
                                }
                            };

                            tryInit();
                        })
                        .catch(err => console.error("Error cargando vista", err));
                }

                function redimWindow(NombreSnippet, ancho) {
                    const padre = "snippet_" + NombreSnippet + "_b";                               
                    document.getElementById(padre).style.width = ancho + "px";
                }

                function centerWindow(nombreSnippet){
                    const win = document.getElementById("snippet_" + nombreSnippet + "_b");

                    win.style.left = (window.innerWidth - win.offsetWidth)/2 + "px";
                    win.style.top  = (window.innerHeight - win.offsetHeight)/2 + "px";
                }

                function bringToFront(elmnt) {
                    let max = 0;

                    document.querySelectorAll(".winBody").forEach(w => {
                        const z = parseInt(w.style.zIndex) || 0;
                        if (z > max) max = z;
                    });

                    elmnt.style.zIndex = max + 10;
                    highestZ = max + 10;
                }                

                function minimizeWindow(nombreSnippet) {
                    const win = document.getElementById("snippet_" + nombreSnippet + "_b");

                    if (!win) return;

                    win.dataset.minimized = "1";
                    win.style.display = "none";

                    const tile = document.createElement("div");
                    tile.className = "taskbar-tile";
                    tile.id = "tile_" + nombreSnippet;
                    tile.innerText = win.dataset.titulo;

                    tile.onclick = function () {
                        restoreWindow(nombreSnippet);
                    };

                    document.getElementById("desk_taskbar").appendChild(tile);
                    
                    navigator.sendBeacon("/core/includes/snip.asp?s=" + nombreSnippet + "&w=1");
                    actualizarTaskbar();
                }

                function maximizeWindow(pagina) {
                    window.location.href = pagina;
                }

                function restoreWindow(nombreSnippet) {
                    const win = document.getElementById("snippet_" + nombreSnippet + "_b");
                    if (!win) return;

                    delete win.dataset.minimized;
                    win.style.display = "block";
                    highestZ += 10;
                    win.style.zIndex = highestZ;

                    const tile = document.getElementById("tile_" + nombreSnippet);
                    if (tile) tile.remove();
                    
                    navigator.sendBeacon("/core/includes/snip.asp?s=" + nombreSnippet + "&w=0");
                    actualizarTaskbar();
                }

                function closeWindow(NombreSnippet) {
                    const closeFunctionName = NombreSnippet + "_close";

                    if (typeof window[closeFunctionName] === "function") {
                        window[closeFunctionName]();
                    }

                    navigator.sendBeacon("/core/includes/snip.asp?s=" + NombreSnippet + "&est=0");
                    const win = document.getElementById("snippet_" + NombreSnippet + "_b");

                    if (win) { 
                        limpiarScriptsVentana(win);
                        win.remove(); 
                    }
                }   

            // Fin - Funciones del Manejador de Ventanas ----------------------------------------------------------------------------

            // Funciones Auxiliares del Manejador de Ventanas y TaskBar -------------------------------------------------------------

                function divFocus(ObjetoHeader, ObjetoBody) {
                    var miDivHeader = document.getElementById(ObjetoHeader.toString());
                    var miDivBody = document.getElementById(ObjetoBody.toString());

                    highestZ += 10;

                    miDivHeader.style.zIndex = highestZ;
                    miDivBody.style.zIndex = highestZ;
                }

                function dragWindow(handle, ventana) {
                    let pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
                    let saveTimeout = null;

                    handle.style.touchAction = "none";

                    handle.addEventListener("pointerdown", function(e) {
                        if (e.target.closest(".window_title_button")) return;

                        e.preventDefault();
                        bringToFront(ventana);
                        pos3 = e.clientX;
                        pos4 = e.clientY;

                        document.addEventListener("pointermove", elementDrag);
                        document.addEventListener("pointerup", closeDragWindow);
                    });

                    function elementDrag(e) {
                        e.preventDefault();

                        pos1 = pos3 - e.clientX;
                        pos2 = pos4 - e.clientY;

                        pos3 = e.clientX;
                        pos4 = e.clientY;

                        ventana.style.top  = (ventana.offsetTop - pos2) + "px";
                        ventana.style.left = (ventana.offsetLeft - pos1) + "px";
                    }

                    function closeDragWindow() {

                        document.removeEventListener("pointerup", closeDragWindow);
                        document.removeEventListener("pointermove", elementDrag);

                        clearTimeout(saveTimeout);

                        saveTimeout = setTimeout(() => {
                            actualizarPersistencia();
                        }, 2500);
                    }
                }     
                
                function limpiarScriptsVentana(div) {
                    if (!div._scripts) return;

                    div._scripts.forEach(s => {
                        if (s && s.parentNode) {
                            s.parentNode.removeChild(s);
                        }
                    });

                    div._scripts = [];
                }   

                function registrarScriptsEnContenedor(container, ownerDiv) {
                    if (!ownerDiv._scripts) ownerDiv._scripts = [];

                    container.querySelectorAll("script").forEach(oldScript => {
                        const newScript = document.createElement("script");

                        if (oldScript.src) {
                            newScript.src = oldScript.src;
                            newScript.defer = false;
                        } else {
                            newScript.textContent = oldScript.textContent;
                        }

                        document.body.appendChild(newScript);
                        ownerDiv._scripts.push(newScript);
                    });
                }                
               
                function win_global(accion) {
                    const ventanas = document
                            .getElementById("desk_div_container")
                            .querySelectorAll(".winBody");

                    ventanas.forEach(win => {
                        const nombre = win.id
                            .replace("snippet_", "")
                            .replace("_b", "");

                        const minimizada = win.dataset.minimized === "1";

                        switch (accion) {
                            case "close":
                                if ( minimizada ) { restoreWindow(nombre); }
                                closeWindow(nombre);
                                break;

                            case "min":
                                if ( !minimizada ) { minimizeWindow(nombre); }
                                break;

                            case "restore":
                                restoreWindow(nombre);
                                break;
                        };
                    });
                }

                function actualizarTaskbar() {
                    const taskbar = document.getElementById("desk_taskbar");

                    if (taskbar.children.length === 0) {
                        taskbar.style.opacity = "0";
                        setTimeout(() => taskbar.style.display = "none", 300);
                    } else {
                        taskbar.style.display = "flex";
                        setTimeout(() => taskbar.style.opacity = "1", 10);
                    }
                }
                
                function actualizarPersistencia() {
                    const ventanas = document
                        .getElementById("desk_div_container")
                        .querySelectorAll(".winBody");

                    let paquete = "";

                    ventanas.forEach(win => {
                        if (win.offsetParent === null) return;
                        if (win.dataset.minimized === "1") return;

                        const nombre = win.id
                            .replace("snippet_", "")
                            .replace("_b", "");

                        const top  = win.offsetTop  || 0;
                        const left = win.offsetLeft || 0;

                        const z = parseInt(window.getComputedStyle(win).zIndex) || 0;

                        paquete += nombre + "|"
                                + top + "|"
                                + left + "|"
                                + z
                                + ";";
                    });

                    if (paquete.length === 0) return;

                    navigator.sendBeacon(
                        "/core/includes/desktop_grabar.asp",
                        paquete
                    );
                }

            // Fin - Funciones Auxiliares del Manejador de Ventanas y TaskBar -------------------------------------------------------            
        </script>  