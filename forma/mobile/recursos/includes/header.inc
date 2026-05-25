        <meta charset="UTF-8" />
        <title><%= Application("Nombre") %></title>

        <link href="/forma/mobile/recursos/images/favicon.png" rel="icon">
        <link href="/forma/mobile/recursos/images/favicon.png" rel="apple-touch-icon">

        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        
        <link rel="stylesheet" href="/forma/mobile/recursos/css/mobile.css" />                
        <link href="/core/lib/font-awesome-6/all.min.css" rel="stylesheet">     

        <link href="/core/lib/swal-2/sweetalert2.min.css" rel="stylesheet">     
        <script src="/core/lib/swal-2/sweetalert2.all.min.js"></script>

        <%
            if not ((request.Cookies("usuario") <> "") and (request.Cookies("nombre") <> "")) then
                response.redirect "/forma/mobile/login.asp"
            end if

            dim mob_conn, mob_table, mob_sqlString, Mob_Vinculo, thisSystem, thisProcess

            set mob_conn = Server.CreateObject("ADODB.Connection")
            mob_conn.open Application("Conn")  

            public Sub SysLockOut()
                dim tt, cmdString, sis, proc

                sis = thisSystem
                proc = thisProcess

                if isnull(Request.Cookies("Usuario")) or (Request.Cookies("Usuario") = "") then 
                    response.redirect "/forma/mobile/login.asp"
                else
                    if isnull(sis) or (sis = "") then sis = "/*"
                    if isnull(proc) or (proc = "") then proc = "/*"          

                    cmdString = "SELECT COUNT(*) AS Acceso " & _
                                "FROM seg_PermisosUsuarios " & _
                                "WHERE (Usuario = '" & Request.Cookies("Usuario") & "') " & _
                                "AND (Sistema = '" & sis & "') " & _
                                "AND (Proceso = '" & proc & "');"

                    set tt = mob_conn.execute(cmdString)
                        if tt("Acceso") = 0 then
                            response.redirect "/forma/mobile/login.asp"
                        end if            
                    tt.close: set tt = nothing
                end if
            end Sub	            

            function NombreUsuario()
                dim cc, tt

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")  
                    set tt = cc.execute("SELECT usuNombre FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("usuario") & "';")
                        NombreUsuario = tt("usuNombre")
                    tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function  

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
        %>

        <script>
            const getVal = id => document.getElementById(id).value;
            const setVal = (id, val) => document.getElementById(id).value = val;
            
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

            function fecha_StrToDate(cadena) {
                const partes = cadena.split('/');
                if (partes.length !== 3) return null;

                const dia = parseInt(partes[0], 10);
                const mes = parseInt(partes[1], 10) - 1; // Los meses van de 0 a 11
                const anio = parseInt(partes[2], 10);

                if (isNaN(dia) || isNaN(mes) || isNaN(anio)) return null;

                return new Date(anio, mes, dia);
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
        </script>