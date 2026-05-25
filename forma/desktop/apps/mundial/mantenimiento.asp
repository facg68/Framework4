<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <title>Polla Mundial</title>
        <meta name="description" content="">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>

        <!-- #include virtual = "/core/includes/kernel/head.inc" -->   
        <%
            thisSystem = "mundial"
            thisProcess = "mundial.050"
            SysLockOut
        %>    
    
        <style>
            body {
                background-color:#9F9F9F;
                background-image:url(Imagenes/fondo.jpg);
                overflow: auto;
            }
			
            td {
                font-family:Verdana, Arial, Helvetica, sans-serif;
            }
			
            #content h2 {
                text-align: left;
            }
			
			.fondo1 { background-color:#E9E9E9; }
			.fondo0 { background-color:#EBF8FE; }	
            
            tr:not(:last-child) { border: none !important; }
        </style>

        <%
			function Acumulado()
				sqlString = "SELECT dbo.mundial_CuantoEnJuego() AS Cuanto;"
							
				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)
				
				if not (t.eof or t.bof) then
					Acumulado = t("Cuanto")
				end if
			
				t.close: set t=nothing
				con.close: set con=nothing
			end function	

			function QueNivel(Etapa)
				sqlString = "SELECT Estatus FROM dbo.mundial_Estatus WHERE Codigo = 'Etapa';"
							
				set con = server.CreateObject("ADODB.Connection")
				con.open Application("Conn")
				set t = con.execute(sqlString)
				
				if not (t.eof or t.bof) then
                    if Etapa = t("Estatus") then
                        QueNivel = "&nbsp;&nbsp;<span style='color: red;'>(actual)</span>"
                    else
                        QueNivel = "&nbsp;"
                    end if
				end if
			
				t.close: set t=nothing
				con.close: set con=nothing
			end function	            		        
        %>
    </head>
	
    <body>
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

		<% Response.Cookies("mundial_key") = "dslkwrtywbfsjbvwegowuienweixmlkjri" %>

        <table style="margin-left: auto; margin-right: auto; width: 95%; padding: 0px; border-spacing: 0px; background-color: rgb(0, 0, 0);">
            <tr style="padding: 0px; border-spacing: 0px;">
                <td colspan="5" style="padding: 0px; border-spacing: 0px;">
                    <img src="Imagenes/header.jpg" style="width: 100%; border-style: none;">
                </td>
            </tr>
        </table>

        <table style="margin-left: auto; margin-right: auto; width: 95%; padding: 0px; border-spacing: 0px; border-style: none;">            
            <tr style="background-color: #F8F8F8;">
                <td style="padding: 30px; font-size: 18px; text-align: left;">
                    Seleccione las opciones de mantenimiento del sistema de Polla mundial

                    <br /><br />

                    Desde esta p&aacute;gina puede realizar las funciones de inicio y final del 
                    juego, por lo que debe estar completamente seguro de seleccionar el proceso 
                    correcto, de otra forma, todo el sistema de Polla Mundial ser&aacute; alterado sin 
                    posibilidad de recuperaci&oacute;n.

                    <br /><br />

                    <span style="font-weight: bold;">Men&uacute; de Opciones:</span>

                    <br /><br />

                    <table style="width:90%;">
                        <tr>
                            <td style="width:5%; text-align: right;">
                                1.&nbsp;
                            </td>

                            <td style="width:95%; font-weight: bold; color: rgb(112, 18, 18);">
                                <div onclick="proceso('Iniciar Sistema', 'man_iniciar.asp')">
                                    Iniciar Sistema
                                </div>
                            </td>
                        </tr>

                        <tr>
                            <td>&nbsp;</td>
                            <td>
                                <br/>
                                Este proceso Inicia la tabla Master, elimina el historial de resultados de los 
                                partidos, borra todas las pollas del sistema, reinicial el nivel de los partidos 
                                a "Fase de Grupos" y cambia los estatus de activaci&oacute;n y creaci&oacute;n de 
                                apuestas.
                                
                                <br/><br/>

                                S&oacute;lo debe usarse ANTES de iniciar el juego. 
                            </td>
                        </tr>

                        <tr><td colspan="2">&nbsp;</td></tr>

                        <tr>
                            <td style="width:5%; text-align: right;">
                                2.&nbsp;
                            </td>

                            <td style="width:95%; font-weight: bold; color: rgb(112, 18, 18);">
                                <div onclick="proceso('Cerrar Periodo de Creacion de Pollas', 'man_estatus_pollas.asp')">
                                    Cerrar el Periodo de Crear Pollas
                                </div>
                            </td>
                        </tr>

                        <tr>
                            <td>&nbsp;</td>
                            <td>
                                <br/>

                                Este proceso Cambia el Estatus del juego y evita que los jugadores 
                                sigan creando Pollas.

                                <br/><br/>

                                Dependiendo de las reglas establecidas, es probable que tambi&eacute;n se 
                                requiera cerrar el proceso de activaci&oacute;n de pollas.
                            </td>
                        </tr>

                        <tr><td colspan="2">&nbsp;</td></tr>

                        <tr>
                            <td style="width:5%; text-align: right;">
                                3.&nbsp;
                            </td>

                            <td style="width:95%; font-weight: bold; color: rgb(112, 18, 18);">
                                <div onclick="proceso('Cerrar Periodo de Activacion de Pollas', 'man_estatus_activacion.asp')">
                                    Cerrar el Proceso de Activaci&oacute;n de Pollas
                                </div>
                            </td>
                        </tr>

                        <tr>
                            <td>&nbsp;</td>
                            <td>
                                <br/>

                                Este proceso Cambia el estatus del juego, de modo que ya no es posible 
                                seguir activando pollas.
                                
                                <br/><br/>

                                De aqu&iacute; en adelante, s&oacute;lo se debe actualizar el Master para 
                                mantener informado a los jugadores acerca de los puntajes que van ganando 
                                sus pollas.
                            </td>
                        </tr>

                        <tr><td colspan="2">&nbsp;</td></tr>

                        <tr>
                            <td style="width:5%; text-align: right;">
                                4.&nbsp;
                            </td>

                            <td style="width:95%; font-weight: bold; color: rgb(112, 18, 18);">
                                <div onclick="proceso('Subir el Nivel de la Etapa del Mundial', 'man_estatus_subir_nivel.asp')">
                                    Subir el Nivel de la Etapa del Juego
                                </div>
                            </td>
                        </tr>

                        <tr>
                            <td>&nbsp;</td>
                            <td>
                                <br/>

                                Este proceso sube en 1 el Nivel de la Etapa del Juego. 
                                Al principio el Nivel es "1" (Fase de Grupos)
                                
                                <br/><br/>

                                De aqu&iacute; en adelante, cada vez que se active esta opcion 
                                el sistema Pasará de un nivel a otro de forma cíclica: <br /><br />

                                0. No Ver Historial de Partidos<%= QueNivel(0) %><br />
                                1. Fase de Grupos<%= QueNivel(1) %><br />
                                2. Octavos de Final<%= QueNivel(2) %><br />
                                3. Cuartos de Final<%= QueNivel(3) %><br />
                                4. Semifinal<%= QueNivel(4) %><br />
                                5. Final<%= QueNivel(5) %><br />
                                6. Ver el Historial Completo de los Partidos<%= QueNivel(6) %><br />
                            </td>
                        </tr>                                    

                        <tr><td colspan="2">&nbsp;</td></tr>     

                        <tr>
                            <td style="width:5%; text-align: right;">
                                5.&nbsp;
                            </td>

                            <td style="width:95%; font-weight: bold; color: rgb(112, 18, 18);">
                                <div onclick="proceso('Finalizar Juego y Calcular Ganadores', 'man_ganadores.asp')">
                                    Finalizar el juego y Calcular Ganadores
                                </div>
                            </td>
                        </tr>

                        <tr>
                            <td>&nbsp;</td>
                            <td>
                                <br/>

                                Esta opci&oacute;n cierra y termina el juego. Se cambia el nivel a "6" (ver todo el historial de 
                                partidos), se cierran el proceso de actualizacion del Master y se proceder&aacute; a calcular a 
                                los ganadores
                                
                                <br/><br/>

                                Este es el proceso final antes de cerrar completamente la polla mundial
                            </td>
                        </tr>

                        <tr><td colspan="2">&nbsp;</td></tr>                                                                                                                     
                    </table>

                    <br/><br/>

                    <div style="text-align: center;">
                        <u><a href="mundial.asp">Pulse aqui para volver a la p&aacute;gina principal.</a></u>
                    </div>

                    <br><br><br/><br/>
                </td>
            </tr>
        </table>

        <script>
            function Requery() {
                document.getElementById("formulario").submit();
            }
                    
            function proceso(Titulo, vinculo) {
                var confirmacion = confirm("Esta completamento seguro de realizar el proceso de " + Titulo + "?");

                if (confirmacion) {
                    window.location.href = vinculo;
                } else {
                    window.alert("Proceso Cancelado.");        
                } 
            }           
        </script>
        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>
</html>