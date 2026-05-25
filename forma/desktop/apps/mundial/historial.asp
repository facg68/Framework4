<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <title>Polla Mundial</title>    
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->
        <%
            thisSystem = "mundial"
            thisProcess = "mundial.035"
            SysLockOut
        %>    

    <style>
      .black {
        color: black;
      }

      td, th {
        padding: 5px;
        font-size: 16px;
      }    

      .ldetalle {
        background-color: rgb(230, 230, 230);
        color: black;        
      }

      .ldetalle_impar {
        background-color: rgb(240, 240, 240);
        color: black;        
      }      

      .top {
        background-color: rgb(71,71,71);
        color: white;
      }

      .foot {
        background-color: rgb(89,89,89);
        color: white;
      }

      .CeldaDetalle {
        border: 1px solid rgb(186, 216, 232);
        padding:5px;
      }

      .vbControl_Enabled {
        background-color: rgb(224, 255, 204);
        color: rgb(0, 0, 0);
        padding: 5px;
        border: 1px solid rgb(199, 199, 199);
      }

      .vbControl_Disabled {
        background-color: rgb(210, 210, 210);
        color: rgb(140, 140, 140);
        padding: 5px;
        border: 1px solid rgb(199, 199, 199);
      }         

      .control-label {
        font-size: 14px;
      }

      .borde {
        border: 1px solid;
        border-color: rgb(184, 184, 184);
      }   

      .gradeA_par {
        background-color: rgb(225, 237, 218);
      }   

      .gradeA_impar {
        background-color: rgb(244, 250, 240);
      } 

      .gradeC_par {
        background-color: rgb(206, 228, 237);
      }   

      .gradeC_impar {
        background-color: rgb(220, 238, 245);
      }     

      .gradeV_par {
        background-color: rgb(201, 201, 201);
      }   

      .gradeV_impar {
        background-color: rgb(224, 224, 224);
      }   

      .res01 { background-color: rgb(255, 255, 255) }
      .res02 { background-color: rgb(231, 240, 216) }  
      
      tr:not(:last-child) { border: none !important; }

      .linea {
        color: rgb(0, 0, 0);
      }
    </style>   

    <%
        dim con, t, tt, sqlString, data, labels
        dim cbox, cuantos, ordenamiento, oo, vv
        dim Codigo, Nombre, Descripcion, Cuenta, vinculo, verTipo   

        set con = Server.CreateObject("ADODB.Connection")
        con.open Application("Conn")     

        Function Usuario_Valido()
          dim con, tfx, sqlString
          
          Usuario_Valido = 0
          sqlString = "exec seg_pa_VerificarPermisoUsuario '" & Request.Cookies("Usuario") & "', 'mundial', 'mundial.035'"

          set tfx = con.execute(sqlString)

          if tfx("Acceso") = 1 then
            Usuario_Valido = 1
          end if
          
          tfx.close: set tfx = nothing
        End Function	

        function HoraNumerica2Hora(Hora)
          dim horas, minutos, sufijo

          sufijo = "A.M."

          if len(Hora) = 4 then
            horas = left(hora, 2)
            minutos = right(hora, 2)

            if cInt(horas) > 12 then
              horas = right("00" & (horas - 12), 2)            
              sufijo = "P.M."
            end if

            HoraNumerica2Hora = cInt(horas) & ":" & minutos & " " & sufijo
          else
            HoraNumerica2Hora = ""
          end if
        end function

        function HoraNumerica2Hora2(Hora)
          dim horas, minutos

          if len(Hora) = 4 then
            horas = left(hora, 2)
            minutos = right(hora, 2)

            HoraNumerica2Hora2 = horas & ":" & minutos
          else
            HoraNumerica2Hora2 = ""
          end if
        end function        

        function FechaNumerica2Fecha(Fecha)
          dim dia, mes, amo

          if len(Fecha) = 8 then
            dia = right(fecha, 2)
            mes = mid(fecha, 5, 2)
            amo = left(fecha, 4)

            FechaNumerica2Fecha = dia & "/" & mes & "/" & amo
          else
            FechaNumerica2Fecha = ""
          end if			
        end function

        function QueEtapa()
          sqlString = "SELECT Estatus FROM dbo.mundial_Estatus WHERE Codigo = 'Etapa';"
                
          set tfx = con.execute(sqlString)
          
          if not (tfx.eof or tfx.bof) then
            QueEtapa = tfx("Estatus")
          end if
        
          tfx.close: set tfx = nothing
        end function		        
    %>       
  </head>
    <body>
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

    <%
        rSec = Request.QueryString("sec")
        vEt = Request.QueryString("e")

        if ( (vEt = "") or (QueEtapa() = "6") ) then 
            vEt = "*"
        else
            vEt = QueEtapa()
        end if        
    %>
    <br />

    <div style="width: 98%; margin: auto;">
      <table style="width: 100%; background-color: rgb(89,89,89); color: white;"> 
        <tr>
          <td>
            <table style="width: 100%; background-color: rgb(89,89,89); color: white;">
              <tr>
                <td style="width: 70%; text-align: left;">
                  <h4>Historial de Partidos</h4>
                </td>

                <td style="width: 30%; text-align: right;">
                  <select name="cboEtapa" id="cboEtapa" class="vbControl_Enabled item" style="width: 50%;" onChange="Requery();">
                    <option value="*" <% if vEt = "*" then response.write "selected" %>>Ver Todo</option>
                    <option value="1" <% if vEt = "1" then response.write "selected" %>>Ver Fase de Grupos</option>
                    <option value="2" <% if vEt = "2" then response.write "selected" %>>Ver Octavos</option>
                    <option value="3" <% if vEt = "3" then response.write "selected" %>>Ver Cuartos</option>
                    <option value="4" <% if vEt = "4" then response.write "selected" %>>Ver Seminfinal</option>
                    <option value="5" <% if vEt = "5" then response.write "selected" %>>Ver Final</option>
                  </select>

                  &nbsp;&nbsp;
                  
                  <button type="button" class="form-btn azul normal" onclick="limpiarCampos();">Nuevo</button>
                </td>                
              </tr>
            </table>

            <table style="width: 100%;"> 
              <tr style="background-color:black; color: white;">
                <td style="width: 15%; text-align: center; padding: 5px;">Etapa</td>
                <td style="width:  5%; text-align: center; padding: 5px;">Grupo</td>

                <td style="width: 15%; text-align: center; padding: 5px;">Fecha</td>
                <td style="width: 25%; text-align: center; padding: 5px;">Equipo 1</td>
                <td style="width:  5%; text-align: center; padding: 5px;">Goles</td>
                <td style="width: 25%; text-align: center; padding: 5px;">Equipo 2</td>
                <td style="width:  5%; text-align: center; padding: 5px;">Goles</td> 

                <td style="width:  5%; text-align: center; padding: 5px;">&nbsp;</td>                                   
              </tr>
            </table>

            <div id="overFlow" style="width:100%; height: 400px; overflow: auto; background-color: rgb(207, 207, 207);">
              <table style="width: 100%;" class="ldetalle borde"> 
                <%
                  set con = Server.CreateObject("ADODB.Connection")
                  con.open Application("Conn")

                  sqlString = "SELECT r.Secuencia, r.Etapa, r.Fecha, r.Hora, r.Grupo, " & _
                                    " r.Equipo1, './Banderas/' + e1.Imagen AS Bandera1, e1.Nombre AS Nombre1, r.Goles1, " & _
                                    " r.Equipo2, './Banderas/' + e2.Imagen AS Bandera2, e2.Nombre AS Nombre2, r.Goles2 " & _
                                "FROM mundial_Resultados AS r " & _
                          "INNER JOIN mundial_Tabla_Equipos AS e2 " & _
                                  "ON r.Equipo2 = e2.Equipo " & _
                          "INNER JOIN mundial_Tabla_Equipos AS e1 " & _
                                  "ON r.Equipo1 = e1.Equipo " 

                  if vEt <> "*" then
                    sqlString = sqlString & " WHERE (r.Etapa = '" & vEt & "') "
                  end if

                  sqlString = sqlString & "ORDER BY r.Etapa DESC, r.Fecha DESC, r.Hora DESC;"

                  set t = con.execute(sqlString)                

                  if not (t.bof or t.eof) then
                      sw = -1
                      cuantos = 0

                      Do
                          sw = -1 * sw 
                          cuantos = cuantos + 1

                          response.write "<tr class='res0"
                            if sw  > 0 then
                              response.write "1"
                            else
                              response.write "2"
                            end if

                            vinculo = "historial.asp?sec=" & t("Secuencia") & "&e=" & t("Etapa")
                          response.write "'>"

                            response.write "<td style='width: 15%; text-align: left;' class='borde'>"
                              response.write "<a class='linea' href='" & vinculo & "'>"
                                Select Case t("Etapa")
                                  Case 1
                                    Response.write "Fase de Grupos"
                                  Case 2
                                    Response.write "Octavos"
                                  Case 3
                                    Response.write "Cuartos"
                                  Case 4
                                    Response.write "Semi Final"
                                  Case 5
                                    Response.write "Final"
                                End Select                            
                              response.write "</a>"
                            response.write "</td>"

                            response.write "<td style='width: 5%; text-align: center;' class='borde'>"
                              response.write "<a class='linea' href='" & vinculo & "'>"
                                response.write t("Grupo")
                              response.write "</a>"
                            response.write "</td>"       

                            response.write "<td style='width: 15%; text-align: center;' class='borde'>"
                              response.write "<a class='linea' href='" & vinculo & "'>"
                                response.write FechaNumerica2Fecha(t("Fecha")) & "<br/>" & HoraNumerica2Hora(t("Hora"))
                              response.write "</a>"
                            response.write "</td>"   

                            response.write "<td style='width: 25; text-align: left;' class='borde'>"
                              response.write "<a class='linea' href='" & vinculo & "'>"

                                bandera_1 = Replace(t("Bandera1"), "Banderas", "imagenes/banderas")

                                if t("Goles1") <> t("Goles2") then
                                  if t("Goles1") > t("Goles2") then
                                    response.write "<img src=" & bandera_1 & " width='33' height='22' border='0'>"
                                    response.write "&nbsp;&nbsp;<span style='color: rgb(3, 50, 168);'>" & t("Nombre1") & "</span>"
                                  else
                                    response.write "<img src=" & Replace(bandera_1, "imagenes/banderas", "imagenes/banderas2") & " width='33' height='22' border='0'>"
                                    response.write "&nbsp;&nbsp;<span style='color: rgb(150, 150, 150);'>" & t("Nombre1") & "</span>"
                                  end if
                                else
                                  response.write "<img src=" & bandera_1 &" width='33' height='22' border='0'>"
                                  response.write "&nbsp;&nbsp;<span style='color: rgb(3, 50, 168);'>" & t("Nombre1") & "</span>"
                                end if

                              response.write "</a>"
                            response.write "</td>" 

                            response.write "<td style='width: 5%; text-align: center;' class='borde'>"
                              response.write "<a class='linea' href='" & vinculo & "'>"

                                if t("Goles1") <> t("Goles2") then
                                  if t("Goles1") > t("Goles2") then
                                    response.write "<span style='color: rgb(3, 50, 168);'>" & t("Goles1") & "</span>"
                                  else
                                    response.write "<span style='color: rgb(150, 150, 150);'>" & t("Goles1") & "</span>"
                                  end if
                                else
                                  response.write "<span style='color: rgb(3, 50, 168);'>" & t("Goles1") & "</span>"
                                end if
                                  
                              response.write "</a>"
                            response.write "</td>"       

                            response.write "<td style='width: 25; text-align: left;' class='borde'>"
                              response.write "<a class='linea' href='" & vinculo & "'>"

                                bandera_2 = Replace(t("Bandera2"), "Banderas", "imagenes/banderas")

                                if t("Goles1") <> t("Goles2") then
                                  if t("Goles2") > t("Goles1") then
                                    response.write "<img src=" & bandera_2 &" width='33' height='22' border='0'>"
                                    response.write "&nbsp;&nbsp;<span style='color: rgb(3, 50, 168);'>" & t("Nombre2") & "</span>"
                                  else
                                    response.write "<img src=" & Replace(bandera_2, "imagenes/banderas", "imagenes/banderas2") & " width='33' height='22' border='0'>"
                                    response.write "&nbsp;&nbsp;<span style='color: rgb(150, 150, 150);'>" & t("Nombre2") & "</span>"
                                  end if
                                else
                                  response.write "<img src=" & bandera_2 &" width='33' height='22' border='0'>"
                                  response.write "&nbsp;&nbsp;<span style='color: rgb(3, 50, 168);'>" & t("Nombre2") & "</span>"
                                end if

                              response.write "</a>"
                            response.write "</td>" 

                            response.write "<td style='width: 5%; text-align: center;' class='borde'>"
                              response.write "<a class='linea' href='" & vinculo & "'>"

                                if t("Goles1") <> t("Goles2") then
                                  if t("Goles2") > t("Goles1") then
                                    response.write "<span style='color: rgb(3, 50, 168);'>" & t("Goles2") & "</span>"
                                  else
                                    response.write "<span style='color: rgb(150, 150, 150);'>" & t("Goles2") & "</span>"
                                  end if
                                else
                                  response.write "<span style='color: rgb(3, 50, 168);'>" & t("Goles2") & "</span>"
                                end if

                              response.write "</a>"
                            response.write "</td>"       

                            response.write "<td style='width: 5%; text-align:center;' class='borde'>"
                                %><a onclick="borrar('<%= t("Secuencia") %>')"><%
                                    response.write "<button class='form-btn rojo'>" 
                                        response.write "<i class=' fa fa-trash fa-xl' title='Borrar Partido'></i>"
                                    response.write "</button>"
                                response.write "</a>"                            
                            response.write "</td>"

                          response.write "</tr>"                        

                          t.MoveNext
                      Loop Until t.eof
                  end if

                  t.close: set t = nothing
                %>
              </table>
            </div>

            <table style="width: 100%;" class="top"> 
                <tr>
                  <td style="font-size: 14px; color: rgb(255, 255, 255); text-align: center; padding: 10px;">
                      <%
                        if cuantos = 0 then
                          response.write "No se ha encontrado ningún partido."
                        else
                          response.write "Se han encontrado " & cuantos & " partidos."
                        end if
                      %>
                  </td>
                </tr>
            </table>

            <table style="width: 100%;" class="top"> 
              <!--
                  Ahora añadimos un formulario para añadir un contacto nuevo
              -->

              <%
                if rSec <> "" then
                  sqlString = "SELECT * FROM mundial_Resultados WHERE Secuencia = " & rSec & ";"

                  set t = con.execute(sqlString)
                    if not (t.bof or t.eof) then
                      Etapa = t("Etapa")
                      Grupo = t("Grupo")
                      Fecha = FechaNumerica2Fecha(t("Fecha"))
                      Hora = HoraNumerica2Hora2(t("Hora"))
                      Equipo1 = t("Equipo1")
                      Goles1 = t("Goles1")
                      Equipo2 = t("Equipo2")
                      Goles2 = t("Goles2")
                      Penales = t("Penales")

                      FechaHora = Fecha & " " & Hora
                    end if
                  t.close: set t = nothing                  
                else
                  Etapa = QueEtapa()
                end if
              %>

              <form name="form_transaccion" id="form_transaccion" method="post" action="historial_grabar.asp">
                <div style="display:none;">
                  <input type="text" name="frmSec" id ="frmSec" value= "<%= rSec %>" />
                </div>

                <%
                    sw = -1 * sw 

                    response.write "<tr class='res0"
                        if (sw > 0) then
                            response.write "1"
                        else
                            response.write "2"
                        end if
                    response.write "'>"            
                %>
                        <td style="width: 12%;" class="borde">
                          <select name="frmEtapa" id="frmEtapa" class="vbControl_Enabled item" style="width: 100%;" >
                            <option value="1" <% if rSec <> "" and Etapa = "1" then response.write "selected" %>>Fase de Grupos</option>
                            <option value="2" <% if rSec <> "" and Etapa = "2" then response.write "selected" %>>Octavos</option>
                            <option value="3" <% if rSec <> "" and Etapa = "3" then response.write "selected" %>>Cuartos</option>
                            <option value="4" <% if rSec <> "" and Etapa = "4" then response.write "selected" %>>Seminfinal</option>
                            <option value="5" <% if rSec <> "" and Etapa = "5" then response.write "selected" %>>Final</option>
                          </select>
                        </td>

                        <td style="width: 5%;" class="borde">
                          <select name="frmGrupo" id="frmGrupo" class="vbControl_Enabled item" style="width: 100%;" >
                            <!--
                                Estos grupos sólo se usan durante la Fase de Grupos
                            -->
                              <option value="A" <% if rSec <> "" and Grupo = "A" then response.write "selected" %>>A</option>
                              <option value="B" <% if rSec <> "" and Grupo = "B" then response.write "selected" %>>B</option>
                              <option value="C" <% if rSec <> "" and Grupo = "C" then response.write "selected" %>>C</option>
                              <option value="D" <% if rSec <> "" and Grupo = "D" then response.write "selected" %>>D</option>
                              <option value="E" <% if rSec <> "" and Grupo = "E" then response.write "selected" %>>E</option>
                              <option value="F" <% if rSec <> "" and Grupo = "F" then response.write "selected" %>>F</option>
                              <option value="G" <% if rSec <> "" and Grupo = "G" then response.write "selected" %>>G</option>
                              <option value="h" <% if rSec <> "" and Grupo = "H" then response.write "selected" %>>H</option>

                            <!--
                                Luego de la Fase de Grupos se usa la etiqueta "-"
                            -->
                              <option value="-" <% if rSec <> "" and Grupo = "-" then response.write "selected" %>>&nbsp;</option>
                          </select>
                        </td>

                        <td style="width: 18%;" class="borde">                        
                          <input class="vbControl_Enabled item" id="frmFechaHora" name="frmFechaHora" type="text" placeholder="dd/mm/aaaa hh:mm" <%
                            if rSec <> "" then
                              response.write " value = '" & FechaHora & "'"
                            end if
                          %> style="width: 100%; font-size: 18px;" />
                        </td>
                        
                        <td style="width: 20%;" class="borde">
                          <select name="frmEquipo1" id="frmEquipo1" class="vbControl_Enabled item" style="width: 100%;" >
                            <%
                              if rSec = "" then
                                sqlString = "SELECT Equipo, Nombre, Imagen, Grupo, EnJuego, Activo " & _
                                            "FROM mundial_Tabla_Equipos " & _
                                            "WHERE (Equipo <> '-') AND (Activo = 1) AND (EnJuego = 1) " & _
                                            "ORDER BY Nombre;"
                              else
                                sqlString = "SELECT Equipo, Nombre, Imagen, Grupo, EnJuego, Activo " & _
                                            "FROM mundial_Tabla_Equipos " & _
                                            "WHERE (Equipo <> '-')  " & _
                                            "ORDER BY Nombre;"
                              end if

                              set t = con.execute(sqlString)
                                if not (t.bof or t.eof) then
                                  Do 
                                    response.write "<option value='" & t("Equipo") & "'"

                                    if rSec <> "" then
                                      if Equipo1 = t("Equipo") then
                                        response.write "selected"
                                      end if
                                    end if

                                    response.write ">" & t("Nombre") & "</option>"

                                    t.MoveNext
                                  Loop until t.eof
                                end if
                              t.close: set t = nothing
                            %>
                          </select>
                        </td>                        

                        <td style="width: 5%;" class="borde">                        
                          <input class="vbControl_Enabled item" id="frmGoles1" name="frmGoles1" type="text" value = "<% 
                            if rSec <> "" then
                              response.write Goles1
                            end if
                          %>" style="width: 100%; font-size: 18px;" />
                        </td>

                        <td style="width: 20%;" class="borde">
                          <select name="frmEquipo2" id="frmEquipo2" class="vbControl_Enabled item" style="width: 100%;" >
                            <%
                              if rSec = "" then
                                sqlString = "SELECT Equipo, Nombre, Imagen, Grupo, EnJuego, Activo " & _
                                            "FROM mundial_Tabla_Equipos " & _
                                            "WHERE (Equipo <> '-') AND (Activo = 1) AND (EnJuego = 1) " & _
                                            "ORDER BY Nombre;"
                              else
                                sqlString = "SELECT Equipo, Nombre, Imagen, Grupo, EnJuego, Activo " & _
                                            "FROM mundial_Tabla_Equipos " & _
                                            "WHERE (Equipo <> '-')  " & _
                                            "ORDER BY Nombre;"
                              end if

                              set t = con.execute(sqlString)
                                if not (t.bof or t.eof) then
                                  Do 
                                    response.write "<option value='" & t("Equipo") & "'"

                                    if rSec <> "" then
                                      if Equipo2 = t("Equipo") then
                                        response.write "selected"
                                      end if
                                    end if                                    

                                    response.write ">" & t("Nombre") & "</option>"

                                    t.MoveNext
                                  Loop until t.eof
                                end if
                              t.close: set t = nothing
                            %>
                          </select>
                        </td>   

                        <td style="width: 5%;" class="borde">                        
                          <input class="vbControl_Enabled item" id="frmGoles2" name="frmGoles2" type="text" value = "<% 
                            if rSec <> "" then
                              response.write Goles2
                            end if
                          %>" style="width: 100%; font-size: 18px;" />
                        </td>

                        <td style="width: 10%;" class="borde">
                          <select name="frmPenales" id="frmPenales" class="vbControl_Enabled item" style="width: 100%;" >
                            <option value="0" <% if rSec <> "" and Penales = "0" then response.write "selected" %>>&nbsp;</option>
                            <option value="1" <% if rSec <> "" and Penales = "1" then response.write "selected" %>>Penales</option>
                          </select>
                        </td>

                        <td style="width: 5%; text-align:center;" class="borde">
                            <button class="form-btn verde" type="submit" style="width: 95%;">
                                <i class=" fa fa-save fa-xl" title="Añadir Nuevo Partido al Historial"></i>
                            </button>
                        </td>
                    </tr>
              </form>
            </table>    
          </td>
        </tr>
      </table>
    </div>

     <script>
        function borrar(secuencia) {
            var confirmacion = confirm("Está seguro de borrar el registro seleccionado?");
            var vinculo = "historial_borrar.asp?s=" + secuencia;

            if (confirmacion) {
                window.location.href = vinculo;
            } else {
                alert("Proceso Cancelado.");        
            }        
        }     

        function limpiarCampos() {
            document.getElementById("frmSec").value = "";
            document.getElementById("frmEtapa").value = "";
            document.getElementById("frmGrupo").value = "";
            document.getElementById("frmFechaHora").value = "";
            document.getElementById("frmEquipo1").value = "";
            document.getElementById("frmGoles1").value = "";
            document.getElementById("frmEquipo2").value = "";
            document.getElementById("frmGoles2").value = "";
        }

        function Requery() {
            var etapa = document.getElementById("cboEtapa").value;
            var vinculo = "historial.asp?e=" + etapa;
            window.location.href = vinculo;
        }     
    </script>

    <script type="text/javascript">
        $(document).ready(function(){
          $('#frmFechaHora').mask('00/00/0000 00:00');
        })
    </script>   

    <%
      con.close: set con = nothing    
    %>
    <!-- #include virtual = "/core/includes/kernel/close.inc" -->
  </body>
</html>