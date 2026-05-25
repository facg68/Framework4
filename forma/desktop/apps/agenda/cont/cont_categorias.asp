<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <title>Editar Categorías de los Contactos</title> 
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->             
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0060"
            SysLockOut


            function limpiar(cadena)
                dim char, k, res

                res = ""

                if len(trim(cadena)) > 0 then
                    for k = 1 to (len(trim(cadena)))
                        char = mid(cadena, k, 1)

                        select case asc(char)
                        case 39: char = "´"
                        case 225: char = "a"
                        case 193: char = "A"
                        case 233: char = "e"
                        case 232: char = "e"
                        case 201: char = "E"
                        case 237: char = "i"
                        case 205: char = "I"
                        case 243: char = "o"
                        case 211: char = "O"
                        case 250: char = "u"
                        case 218: char = "U"
                        case 209: char = "N"
                        case 241: char = "n"
                        end select

                        res = res  & char
                    next
                end if

                limpiar = res
            end function   

            function NombreTipo(Tipo)
                dim cc, tt, sqlString

                sqlString = "SELECT Nombre FROM con_Contactos_Tipos WHERE Codigo = '" &  Tipo & "';"

                set cc = Server.CreateObject("ADODB.Connection")
                cc.open Application("Conn")
                    set tt = cc.execute(sqlString)
                        NombreTipo = tt("Nombre")
                    tt.close: set tt = nothing
                cc.close: set cc = nothing
            end function
        %> 

        <style>
            .fila {
                display: flex;
                align-items: center;                               
                gap: 10px;
            }

            .col1 {
                white-space: nowrap;
                font-weight: bold;
                min-width: fit-content;
            }

            .col2 { flex: 1; }
        </style>           
    </head>

    <body plantilla="tabla" reserva="225">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            dim con, t, sqlString, cuantos, usu
            dim Codigo, Nombre, Vacio, Tipo

            cuantos = 0
            Tipo = Request.QueryString("t")
            usu = Request.Cookies("Usuario")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            sqlString = "SELECT Codigo, Tipo, Nombre, PorDefecto, DeSistema, dbo.ContarContactosCateg(Usuario, Tipo, Codigo) AS Cuantos " & _
                        "FROM dbo.con_Contactos_Categorias " & _
                        "WHERE (Usuario = '" & usu & "') " & _ 
                        "AND (Tipo = '" & Tipo & "') " & _
                        "ORDER BY Nombre;"
    
            set t = con.execute(sqlString)
        %>        

        <br />

        <div style="display: flex; justify-content: space-between; width: 92%; margin: auto;">
            <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                Categorias del Tipo <%= NombreTipo(Tipo)%>
            </div>
            
            <div style="flex: 0 0 50%; text-align: right;">
                <button type="button" class="form-btn normal azul" onclick="volver()">
                    Volver
                </button>  
            </div>
        </div>        

        <div class="main" style="width: 95%;">
            <div class="line">
                <div class="tabla-wrapper">
                    <table class="tabla tabla-carbon">
                        <thead>
                            <tr>
                                <th class="sticky" style="width: 85%;">Nombre</th>
                                <th class="sticky" style="width: 15%; text-align: center;">Acciones</th>
                            </tr>
                        </thead>

                        <tbody>                    
                            <%
                                if not (t.bof or t.eof) then
                                    cuantos = 0

                                    Do
                                        cuantos = cuantos + 1 
                                        subClase = "tr-verde"
                                        estado = ""

                                        if t("Cuantos") > 0 then subClase = "tr-rojo"      ' Tiene información - No se puede borrar
                                        if t("DeSistema") = 1 then subClase = "tr-azul"    ' Es de Sistema - No puede tocarse

                                        response.write "<tr class='" & subClase & "'>"
                                            response.write "<td>" & limpiar(t("Nombre")) & "</td>"

                                            response.write "<td style='text-align: center;'>"
                                                if (subClase <> "tr-azul") then
                                                    if (subClase = "tr-rojo" ) then estado =  " disabled"

                                                    %>
                                                        <button class = "form-btn rojo <%= estado %>" 
                                                                type = "button" 
                                                                onclick="borrar('<%= Tipo %>', '<%= t("Codigo") %>', '<%= t("Nombre") %>', '<%= t("Cuantos") %>')" 
                                                                <%= estado %>>
                                                            <i class=' fa fa-trash fa-xl' title='Borrar Categoria'></i>
                                                        </button>
                                                    <%
                                                else
                                                    %>
                                                        <button class = "form-btn rojo" 
                                                                style = "visibility: hidden;"
                                                                type = "button">                                                                
                                                            <i class=' fa fa-trash fa-xl' title='Borrar Categoria'></i>
                                                        </button>
                                                    <%                                                
                                                end if
                                            response.write "</td>"
                                        response.write "</tr>"                        

                                        t.MoveNext
                                    Loop Until t.eof
                                end if
                            %>
                        </tbody>

                        <tfoot>
                            <tr>
                                <td class="sticky" style="text-align: center;" colspan="2">
                                    <%
                                        response.write "Se encontraron " & cuantos & " Categorías de Contactos"
                                    %>
                                </td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>

            <form name="formulario" id="formulario" method="post" action="cont_categorias_nueva_cat.asp">
                <div class="no-ver">
                    <input id="codTipo" name="codTipo" type="text" value="<%= Tipo %>"/>                
                </div>

                <div class="fila">
                    <div class="col col1">Nueva Categoría</div>

                    <div class="col col2">
                        <input class="field" style="width: 100%;" type="text" id="nuevoNombre" name="nuevoNombre" >
                    </div>

                    <div class="col col3">
                        <button class="form-btn verde " type="submit">
                            <i class="fa fa-save fa-xl" title="Añadir"></i>
                        </button>   
                    </div>
                </div>              
            </form>
        </div>
  
        <br /><br />   

        <script type="text/javascript">
            function borrar(tipo, codigo, nombre, cuantos) {
                if (cuantos == 0) {
                    var confirmacion = confirm("Está seguro de borrar la categoria " + nombre + "?");
                    var vinculo = "cont_categorias_borrar_cat.asp?t=<%= Tipo %>&c=" + codigo;

                    if (confirmacion) {
                        window.location.href = vinculo;
                    } else {
                        alert("Proceso Cancelado.");
                    }        
                }
                else {
                    window.alert("Esta categoria ya ha sido asignada a, por lo menos, un contacto. No puede ser eliminada!"); 
                }
            }   
            
            function volver() {
                window.location.href = "cont_tipos.asp";
            }
        </script>

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>