<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<%
    '---------------------------------------------------------------'
    '                                                               '
    ' ESTE ES UN PROCESO DESTRUCTIVO. AL MODIFICAR ROLES, ESTOS SON '
    ' ELIMINADOS DEL SISTEMA Y SE VUELVEN A CONSTRUIR USANDO LOS    '
    ' VALORES DEL FORMULARIO. SI SUCEDE ALGUN ERROR, EL  ROL QUEDA  '
    ' COMPLETAMENTE DESTRUIDO.                                      '
    '                                                               '
    '---------------------------------------------------------------'

    '
    ' Init()
    '

    dim cc, s, p, sqlString, NombreCampo, ValorCampo, Codigo, Nombre, Descripcion, TipoRol, OrdenadoPor, unico, vinculo

    set cc = Server.CreateObject("ADODB.Connection")
    cc.open Application("Conn")

    '
    ' Funciones y Procedimientos
    '

    function limpiar(cadena)
        limpiar = Replace(cadena,"'","´")    
    end function

    function CuantosSistemas(Rol)
        dim tt, sqlCommand

        sqlCommand = "SELECT COUNT(detRolSistema) AS Cuantos " & _
                     "FROM (SELECT DISTINCT detRolSistema " & _
                     "FROM dbo.seg_RolDetalles " & _
                     "WHERE (detRol = '" & Rol & "')) AS t;"

        set tt = cc.execute(sqlCommand)
            CuantosSistemas = tt("Cuantos")
        tt.close: set tt = nothing    
    end function

    function QueSistema(Rol)
        dim tt, sqlCommand

        sqlCommand = "SELECT DISTINCT detRolSistema " & _
                     "FROM dbo.seg_RolDetalles " & _
                     "WHERE (detRol = '" & Rol & "');"

        set tt = cc.execute(sqlCommand)
            QueSistema = tt("detRolSistema")
        tt.close: set tt = nothing
    end function


    '
    ' Main()
    '

    if Request.Cookies("usuario") = "" then
        Response.Redirect "/apps/desktop/login.asp"      
    end if

    Nuevo = Request.Form("Nuevo")

    Codigo = Request.Form("Codigo")
    Nombre = limpiar(Request.Form("Nombre"))
    Descripcion = limpiar(Request.Form("Descripcion"))
    TipoRol = Request.Form("TipoRol")
    OrdenadoPor = Request.Form("OrdenadoPor")
    unico = Request.Form("unico")

    '
    ' Parte 1: Actualizar o Crear Rol
    '

    if Nuevo = "1" then
        Codigo = limpiar(Codigo)

        sqlString = "INSERT INTO seg_Roles(rolCodigo, rolNombre, rolDescripcion, TipoRol) " & _
                            "VALUES ('" & Codigo & "', '" & Nombre & "', '" & Descripcion & "', " & TipoRol & ");"
    else
        sqlString = "UPDATE seg_Roles " & _
                        "SET rolNombre = '" & Nombre & "'," & _
                            " rolDescripcion = '" & Descripcion & "'," & _
                            " TipoRol = " & TipoRol & _
                    " WHERE (rolCodigo = '" & Codigo & "');"
    end if

    cc.execute(sqlString)

response.write sqlString & "<br/><br/><br/><br/>"


    '
    ' Parte 2: Actualizar Procesos del Rol
    '
    set s = cc.execute("SELECT sysCodigo FROM seg_Sistemas;")
        if not (s.bof or s.eof) then
            '
            ' BORRAMOS los detalles del Rol...
            '
            cc.execute("DELETE FROM seg_RoLDetalles WHERE detRol = '" & Codigo & "';")        

            Do
                set p = cc.execute("SELECT proSistema, proCodigo FROM seg_Procesos WHERE proSistema = '" & s("sysCodigo") & "';")
                    if not (p.bof or p.eof) then
                        Do
                            NombreCampo = p("proSistema") & "__" & p("proCodigo") 
                            NombreCampo = Replace(NombreCampo, ".", "_")

                            ValorCampo = Request.Form(NombreCampo)
                            if ValorCampo = 1 then
                                '
                                ' El campo está "Checked"
                                '
                                sqlString = "INSERT INTO seg_rolDetalles(detRol, detRolSistema, detRolProceso) " & _
                                                 "VALUES ('" & Codigo & "','" & p("proSistema") & "','" & p("proCodigo") & "');"

                                response.write NombreCampo & "(" & ValorCampo & ")    -   " & sqlString & "<br/>"  
                                cc.execute(sqlString)                              
                            end if

                            p.MoveNext
                        Loop Until p.eof
                    end if

                s.MoveNext
            Loop Until s.eof
        end if
    s.close: set s = nothing

    '
    ' Parte 3: Actualizamos el Sistema del Rol si todos 
    '          los procesos son del mismo sistema
    '

    if CuantosSistemas(Codigo) = 1 then
        sqlString = "UPDATE seg_Roles " & _
                        "SET CodigoSys = '" & QueSistema(Codigo) & "' " & _
                    " WHERE (rolCodigo = '" & Codigo & "');"

        cc.execute(sqlString)   
    else         
        sqlString = "UPDATE seg_Roles " & _
                        "SET CodigoSys = NULL " & _
                    " WHERE (rolCodigo = '" & Codigo & "');"

        cc.execute(sqlString)       
    end if

    '
    ' Fin del Proceso
    '

    if unico = "1" then
        vinculo = "roles_sys.asp?s=" & QueSistema(Codigo) & "&o=" & OrdenadoPor
    else
        vinculo = "roles.asp?o=" & OrdenadoPor
    end if

    cc.close: set cc = nothing

    response.redirect vinculo
%>