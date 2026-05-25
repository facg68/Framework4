<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <title>Posible Ataque Detectado!</title>
    </head>
    
    <body>
    	<%
			dim token, tipo, cuantos
			
			token = request.QueryString("k")
			tipo = request.QueryString("t")
			
			response.write "<br /><br />"
			response.write "Se ha detectado un posible ataque del tipo SQL Injection en la Intranet"
		%>

		<hr>
        
        El ataque se dio en el valor [&nbsp; <%= token %> &nbsp;].<br />
        El tipo de datos por donde entro el ataque es <%= tipo %>.

        <br /><br />
        El Proceso se ha Detenido. Presione "Volver" en el navegador para que el sistema intente 
        volver a donde se encontraba al ser detectado el problema.
    </body>
</html>
