<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
    <head>
        <title><%= Application("Nombre") %></title>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>   
        <meta name="viewport" content="width=device-width, initial-scale=1.0">             

        <!-- Íconos para navegadores normales / Android -->
        <link rel="icon" type="image/png" sizes="32x32" href="/core/imagenes/favicon.png" />
       
        <!-- Íconos específicos para iOS (Web Clips) -->
        <link rel="apple-touch-icon" sizes="180x180" href="/core/imagenes/icons/apple_180.png">
        <link rel="apple-touch-icon" sizes="152x152" href="/core/imagenes/icons/apple_152.png">
        <link rel="apple-touch-icon" sizes="120x120" href="/core/imagenes/icons/apple_120.png">
        <link rel="apple-touch-icon" sizes="76x76"   href="/core/imagenes/icons/apple_76.png">

        <!-- Meta tags opcionales para iOS -->
        <meta name="apple-mobile-web-app-capable" content="yes">
        <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">        

        <style type="text/css">
            body, html { margin: 0; padding: 0; height: 100%; overflow: hidden; background-color: black; }
            iframe { overflow: hidden; border: none; width: 100%; height: 100%; }            
            #content { position:absolute; left: 0; right: 0; bottom: 0; top: 0px; }
        </style>
    </head>

    <body>
        <div id="content">
            <iframe 
                width="100%" 
                height="100%" 
                frameborder="0" 
                src="pc_mobile.asp" 
                style="overflow: hidden;">
            </iframe>
        </div>        
    </body>
</html>