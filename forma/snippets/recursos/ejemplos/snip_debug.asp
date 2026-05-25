<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">

        <!-- #include virtual = "/core/includes/kernel/head.inc" -->    

        <style>
            html, body{
                height:100%;
                margin:0;
                padding:0;
            }

            .desktop_container_menu{
                height:100vh;
                position:absolute;
                top:65px;
                right:0;
                left:0;
                bottom:0;
                z-index:1;
            }

            #desk_taskbar{
                position:fixed;
                bottom:0;
                left:0;
                right:0;

                height:34px;

                transition:opacity .3s ease;
                opacity:0;
                display:none;

                align-items:center;
                gap:4px;
                padding:4px;

                background:linear-gradient(to top, rgba(40,40,40,.45), rgba(80,80,80,.15) );

                backdrop-filter:blur(6px);
                -webkit-backdrop-filter:blur(6px);

                z-index:9999;
            }

            .taskbar-tile{
                background:hsla(0,0%,28%,.13);
                border:1px solid rgba(255,255,255,.06);

                color:#fff;

                padding:3px 12px;

                font:12px sans-serif;

                border-radius:6px;

                backdrop-filter:blur(4px);
                -webkit-backdrop-filter:blur(4px);

                cursor:pointer;
            }

            .taskbar-tile:hover{
                background:rgb(206,206,206);
                color:black;
            }

            #debug_console{
                position:fixed;
                right:0;
                bottom:34px;

                width:420px;
                max-height:40vh;

                overflow-y:auto;

                background:rgba(0,0,0,.9);

                color:#00ff9c;

                font:12px monospace;

                padding:10px;

                border-left:1px solid rgba(255,255,255,.2);
                border-top:1px solid rgba(255,255,255,.2);

                z-index:99999;
            }

            #debug_console button{
                margin-bottom:8px;
                font-size:11px;
            }
        </style>

        <%
            dim Snippet
            Snippet = Request.QueryString("snippet")
        %>
    </head>

    <body class="bg" plantilla="desktop" reserva="0"
          style="background-color:rgb(77,109,143);"
          onload="init()">

        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <div id="desk_div_container" style="overflow:hidden;" class="desktop_container_menu">
        </div>

        <div id="desk_taskbar"></div>

        <div id="debug_console">
            <button onclick="clearDebug()">Limpiar consola</button>
        </div>

        <script>
            var highestZ;

            function debugLog(msg,type="info"){
                const consoleDiv=document.getElementById("debug_console");
                const line=document.createElement("div");

                if(type==="error")line.style.color="#ff6b6b";
                if(type==="warn")line.style.color="#ffd166";

                line.textContent="["+type.toUpperCase()+"] "+msg;
                consoleDiv.appendChild(line);
                consoleDiv.scrollTop=consoleDiv.scrollHeight;
            }

            function clearDebug(){
                document.getElementById("debug_console").innerHTML=
                '<button onclick="clearDebug()">Limpiar consola</button>';
            }

            window.onerror=function(msg,url,line,col,error){
                debugLog(msg+" ("+url+":"+line+":"+col+")","error");
                return false;
            };

            window.addEventListener("unhandledrejection",function(e){
                debugLog("Promise error: "+e.reason,"error");
            });

            function init(){
                highestZ=20;
                debugLog("Debugger iniciado");

                document.addEventListener("keydown",function(e){
                    if(e.ctrlKey && e.key==="d"){
                        const snip=prompt("Snippet a abrir:");

                        if(snip){
                            openWindow(snip,50,50,snip,"",0);
                        }
                    }
                });

                openWindow('<%= Snippet %>',50,50,"Debugger","",0);
            }

            function divFocus(ObjetoHeader,ObjetoBody){
                var miDivHeader=document.getElementById(ObjetoHeader.toString());
                var miDivBody=document.getElementById(ObjetoBody.toString());

                highestZ+=10;

                miDivHeader.style.zIndex=highestZ;
                miDivBody.style.zIndex=highestZ;
            }

            function openWindow(nombreObjeto,top,left,titulo,accion,minimizado=0){
                debugLog("Cargando snippet: "+nombreObjeto);
                const existing=document.getElementById("snippet_"+nombreObjeto+"_b");

                if(existing){
                    highestZ+=10;
                    existing.style.zIndex=highestZ;
                    return;
                }

                const div=document.createElement("div");

                highestZ+=10;

                div.id="snippet_"+nombreObjeto+"_b";
                div.className="winBody";

                div.style.width="0px";
                div.style.left=left+"px";
                div.style.top=top+"px";

                div.style.display="none";
                div.style.position="absolute";

                div.style.zIndex=highestZ;

                div.innerHTML=`<div class="winHeader"
                                    id="snippet_${nombreObjeto}_h"
                                    onclick="divFocus('snippet_${nombreObjeto}_b','snippet_${nombreObjeto}_h')"
                                    style="display:flex;align-items:center;
                                    background-color:var(--menu-barra-fondo);
                                    color:var(--menu-barra-font-color);">

                                    <div style="flex:1;text-align:left;overflow:hidden;
                                                white-space:nowrap;text-overflow:ellipsis;
                                                font:<%= framework_menu_font_size() %>px <%= framework_mb_font() %>;
                                                color:<%= framework_mb_font_color() %>;">
                                        &nbsp;${titulo}
                                    </div>

                                    <div style="margin-left:auto;display:flex;gap:6px;align-items:center;">
                                        <img class="clickable window_title_button"
                                             src="/core/imagenes/min.png"
                                             onclick="event.stopPropagation();minimizeWindow('${nombreObjeto}','${titulo}')"/>

                                        <img class="clickable window_title_button"
                                             src="/core/imagenes/max.png"
                                             onclick="event.stopPropagation();maximizeWindow('${accion}')"/>

                                        <img class="clickable window_title_button"
                                             src="/core/imagenes/cerrar.png"
                                             onclick="event.stopPropagation();closeWindow('${nombreObjeto}')"/>

                                        &nbsp;
                                    </div>
                                </div>

                                <div id="snippet_${nombreObjeto}_c"></div>
                    `;

                document.getElementById("desk_div_container").appendChild(div);

                const header=div.querySelector(".winHeader");
                dragWindow(header,div);
                navigator.sendBeacon("/core/includes/snip.asp?s="+nombreObjeto+"&est=1");

                fetch("/forma/snippets/"+nombreObjeto+".asp")
                .then(r=>{
                    if(!r.ok){
                        throw new Error("HTTP "+r.status);
                    }
                    return r.text();
                })

                .then(html=>{
                    const body=document.getElementById("snippet_"+nombreObjeto+"_c");
                    body.innerHTML=html;
                    body.querySelectorAll("script").forEach(oldScript=>{
                        try{
                            const newScript=document.createElement("script");
                            newScript.textContent=oldScript.textContent;
                            document.body.appendChild(newScript);
                        }catch(err){
                            debugLog("Script error: "+err.message,"error");
                        }
                });

                const initFunctionName=nombreObjeto+"_init";
                let tries=0;

                const tryInit=()=>{
                    if(typeof window[initFunctionName]==="function"){
                        debugLog("Init encontrado: "+initFunctionName);
                        window[initFunctionName](div);

                        reemplazarEstiloPorClase("window_title_button", filter("<%= icon_foreground_color %>"));

                        if(minimizado==1){
                            minimizeWindow(nombreObjeto,titulo);
                        }else{
                            div.style.display="block";
                        }
                    }else{
                        tries++;
                        if(tries>200){
                            debugLog("Init no encontrado: "+initFunctionName,"error");
                            return;
                        }

                        setTimeout(tryInit,10);
                    }
                };

                tryInit();
                    debugLog("Snippet cargado: "+nombreObjeto);
                })
                .catch(err=>{
                    console.error(err);
                    debugLog("Error cargando "+nombreObjeto+": "+err.message,"error");
                    const body=document.getElementById("snippet_"+nombreObjeto+"_c");

                    if(body){
                        body.innerHTML=
                            '<div style="color:red;padding:20px;font-family:monospace">'+
                            'ERROR cargando snippet<br><br>'+
                            err.message+
                            '</div>';
                    }
                });
            }

            function loadInWindow(nombreSnippet, url) {
                if (!url || typeof url !== "string") {
                    console.warn("URL inválida");
                    return;
                }

                const body = document.getElementById("snippet_" + nombreSnippet + "_c");
                const win  = document.getElementById("snippet_" + nombreSnippet + "_b");

                if (!body || !win) return;

                fetch(url)
                    .then(r => r.text())
                    .then(html => {

                        body.innerHTML = html;

                        body.querySelectorAll("script").forEach(oldScript => {
                            const newScript = document.createElement("script");
                            newScript.textContent = oldScript.textContent;
                            document.body.appendChild(newScript);
                        });

                        const initFunctionName = nombreSnippet + "_init";

                        const tryInit = () => {
                            if (typeof window[initFunctionName] === "function") {
                                window[initFunctionName](win);

                                reemplazarEstiloPorClase(
                                    "window_title_button",
                                    filter("<%= icon_foreground_color %>")
                                );
                            } else {
                                setTimeout(tryInit, 10);
                            }
                        };

                        tryInit();
                    })
                    .catch(err => console.error("Error cargando vista", err));
            }            

            function closeWindow(NombreSnippet){
                const closeFunctionName=NombreSnippet+"_close";

                if(typeof window[closeFunctionName]==="function"){
                    window[closeFunctionName]();
                }

                navigator.sendBeacon("/core/includes/snip.asp?s="+NombreSnippet+"&est=0");
                const win=document.getElementById("snippet_"+NombreSnippet+"_b");

                if(win){win.remove();}
            }

            function redimWindow(NombreSnippet, ancho) {
                const padre = "snippet_" + NombreSnippet + "_b";                               
                document.getElementById(padre).style.width = ancho + "px";
            }            

            function minimizeWindow(nombreSnippet,titulo){
                const win=document.getElementById("snippet_"+nombreSnippet+"_b");

                if(!win)return;

                win.dataset.minimized="1";
                win.style.display="none";

                const tile=document.createElement("div");

                tile.className="taskbar-tile";
                tile.id="tile_"+nombreSnippet;
                tile.innerText=titulo;
                tile.onclick=function(){
                    restoreWindow(nombreSnippet);
                };

                document.getElementById("desk_taskbar").appendChild(tile);
                navigator.sendBeacon("/core/includes/snip.asp?s="+nombreSnippet+"&w=1");
                actualizarTaskbar();
            }

            function restoreWindow(nombreSnippet){
                const win=document.getElementById("snippet_"+nombreSnippet+"_b");

                if(!win)return;
                delete win.dataset.minimized;

                win.style.display="block";
                highestZ+=10;
                win.style.zIndex=highestZ;

                const tile=document.getElementById("tile_"+nombreSnippet);

                if(tile)tile.remove();
                navigator.sendBeacon("/core/includes/snip.asp?s="+nombreSnippet+"&w=0");
                actualizarTaskbar();
            }

            function maximizeWindow(pagina){
                window.location.href=pagina;
            }

            function actualizarTaskbar(){
                const taskbar=document.getElementById("desk_taskbar");

                if(taskbar.children.length===0){
                    taskbar.style.opacity="0";

                    setTimeout(()=>taskbar.style.display="none",300);
                }else{
                    taskbar.style.display="flex";
                    setTimeout(()=>taskbar.style.opacity="1",10);
                }
            }

            function bringToFront(elmnt){
                let max=0;

                document.querySelectorAll(".winBody").forEach(w=>{
                    const z=parseInt(w.style.zIndex)||0;
                    if(z>max)max=z;

                });

                elmnt.style.zIndex=max+10;
                highestZ=max+10;
            }

            function dragWindow(handle,ventana){
                let pos1=0,pos2=0,pos3=0,pos4=0;

                handle.style.touchAction="none";
                handle.addEventListener("pointerdown",function(e){
                    if(e.target.closest(".window_title_button"))return;

                    e.preventDefault();
                    bringToFront(ventana);

                    pos3=e.clientX;
                    pos4=e.clientY;

                    document.addEventListener("pointermove",elementDrag);
                    document.addEventListener("pointerup",closeDragWindow);

                });

                function elementDrag(e){
                    e.preventDefault();

                    pos1=pos3-e.clientX;
                    pos2=pos4-e.clientY;

                    pos3=e.clientX;
                    pos4=e.clientY;

                    ventana.style.top=(ventana.offsetTop-pos2)+"px";
                    ventana.style.left=(ventana.offsetLeft-pos1)+"px";
                }

                function closeDragWindow(){
                    document.removeEventListener("pointerup",closeDragWindow);
                    document.removeEventListener("pointermove",elementDrag);
                }
            }
        </script>
    </body>

    <!-- #include virtual = "/core/includes/kernel/close.inc" -->
</html>