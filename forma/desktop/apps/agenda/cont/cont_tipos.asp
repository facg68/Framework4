<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <title>Editar Tipos de Contactos</title>    
        <!-- #include virtual = "/core/includes/kernel/head.inc" -->             
        <%
            thisSystem = "agenda"
            thisProcess = "agenda.0060"
            SysLockOut
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

    <body plantilla="tabla" reserva="215">
        <!-- #include virtual = "/core/includes/kernel/body.inc" -->

        <%
            dim con, t, sqlString, cuantos
            dim Codigo, Nombre, Vacio, usu

            cuantos = 0
            usu = Request.Cookies("Usuario")

            set con = Server.CreateObject("ADODB.Connection")
            con.open Application("Conn")

            sqlString = "SELECT Codigo, Nombre, DeSistema, ISNULL(Cuantos, 0) AS Cuantos " & _
                        "FROM con_Contactos_Tipos AS t " & _
                        "LEFT JOIN ( " & _
                                        "SELECT Tipo, count(*) AS Cuantos " & _
                                        "FROM con_Contactos_Categorias " & _
                                        "WHERE (Usuario = '" & usu & "') " & _
                                        "GROUP BY Tipo " & _
                                    ") AS q " & _
                        "ON t.Codigo = q.Tipo " & _
                        "WHERE (t.Usuario = '" & usu & "') " & _
                        "ORDER BY t.Nombre ASC;"

            set t = con.execute(sqlString)
        %>          

        <br />

        <div style="display: flex; justify-content: space-between; width: 95%; margin: auto;">
            <div style="flex: 0 0 50%; text-align: left; font-size: 25px; color: rgb(50, 50, 50);">
                Tipos de Contactos
            </div>
            
            <div style="flex: 0 0 50%; text-align: right;">
                &nbsp;
            </div>
        </div>        

        <div class="main" style="width: 95%;">
            <div class="line">
                <div class="tabla-wrapper">
                    <table class="tabla tabla-carbon">
                        <thead>
                            <tr>
                                <th class="sticky" style="width: 85%;">Nombre</th>
                                <th class="sticky" style="width: 15%;">&nbsp</th>
                            </tr>
                        </thead>

                        <tbody>                    
                            <%
                                if not (t.bof or t.eof) then
                                    Cuantos = 0

                                    Do
                                        Cuantos = Cuantos + 1
                                        subClase = "tr-verde"
                                        estado = ""

                                        if t("Cuantos") > 0 then subClase = "tr-rojo"      ' Tiene información - No se puede borrar
                                        if t("DeSistema") = 1 then subClase = "tr-azul"    ' Es de Sistema - No puede tocarse                                        

                                        response.write "<tr class='" & subClase & "'>"
                                            response.write "<td>" & t("Nombre") & "</td>"

                                            response.write "<td style='text-align: center;'>"
                                                if (subClase <> "tr-azul") then
                                                    if (subClase = "tr-rojo" ) then estado =  " disabled"

                                                    %>
                                                        <button class = 'form-btn azul'
                                                                type = "button" 
                                                                onclick = "categorias('<%= t("Codigo") %>')">
                                                            <i class=' fa fa-edit fa-xl' title='Editar Categorías'></i>
                                                        </button>

                                                        <button class = 'form-btn rojo <%= estado %>'
                                                                type = "button" 
                                                                onclick = "borrar('<%= usu %>', '<%= t("Codigo") %>', '<%= t("Nombre") %>')" 
                                                                <%= estado %>>
                                                            <i class=' fa fa-trash fa-xl' title='Borrar Tipo'></i>
                                                        </button>
                                                    <%   
                                                else
                                                    %>
                                                        <button class = 'form-btn azul'
                                                                type = "button" 
                                                                onclick = "categorias('<%= t("Codigo") %>')">
                                                            <i class=' fa fa-edit fa-xl' title='Editar Categorías'></i>
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
                                        response.write "Se encontraron " & cuantos & " Tipos de Contactos"
                                    %>
                                </td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>

            <form name="formulario" id="formulario" method="post" action="cont_nuevo_tipo.asp">
                <div class="fila">
                    <div class="col col1">Nuevo Tipo</div>

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
            function borrar(usu, codigo, nombre) {
                var confirmacion = confirm("Está seguro de borrar el Tipo " + nombre + "?");
                var vinculo = "cont_borrar_tipo.asp?u=" + usu + "&t=" + codigo;

                if (confirmacion) {
                    window.location.href = vinculo;
                } else {
                    alert("Proceso Cancelado.");
                }        
            }   
            
            function categorias(codigo) {
                var vinculo = "cont_categorias.asp?t=" + codigo;
                window.location.href = vinculo;
            }               
        </script>

        <!-- #include virtual = "/core/includes/kernel/close.inc" -->
    </body>