<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <%
            dim cc, tt, sqlString, cuantos, inicio, upaq
            dim fSysObj, imgObjeto, Jpeg, uPath
            dim original, nueva, Path

            set fSysObj = Server.CreateObject("Scripting.FileSystemObject")            
            Set Jpeg = Server.CreateObject("Persits.Jpeg") 
            uPath = "E:\web\facg68co\perfiles\fcardenas\fotos"
            vPath = lcase(Request.Cookies("usuPath")) & "/fotos"

            sub Resize(contacto)
                Original = contacto & ".jpg" 
                Nueva = contacto & "_s.jpg" 

                Path = Server.MapPath(vPath & "/" & Original) 
                Jpeg.Open Path 

                Jpeg.Width = 60
                Jpeg.Height = cInt((60 * Jpeg.OriginalHeight) / Jpeg.OriginalWidth)

                Jpeg.Save Server.MapPath(vPath & "/" & Nueva)         
            end sub
        %>
    </head>

    <body>
        <% 
            set cc = Server.CreateObject("ADODB.Connection")
            cc.Open Application("Conn")

            '-------------------------------
            '                               '
            ' Parte 1: Lista de Contactos   '
            '                               '
            '--------------------------------

            sqlString = "SELECT Codigo from con_Contactos ORDER BY Codigo ASC;"

            set tt = cc.execute(sqlString)

            if not (tt.bof or tt.eof) then
                inicio = 0
                Do  

                    imgObjeto =  uPath & "\" & tt("Codigo") & ".jpg" 
    response.write "Revisando " & imgObjeto & "..."
                    if fSysObj.FileExists(imgObjeto) then 
    response.write "Existe<br/>"
                        upaq = tt("Codigo")
                        Resize upaq
                    else
    response.write "<br/>"
                    end if

                    tt.MoveNext                        
                Loop Until (tt.eof) 

                response.write "Se revisaron " & inicio & " paquetes -- " & upaq
            end if

            tt.close: set tt = nothing
            cc.close: set cc = nothing    

            set Jpeg = nothing
            set fSysObj = nothing        
        %>
    </body>
</html>