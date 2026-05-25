<%@ CodePage=65001 %>
<!-- #include virtual = "/core/includes/kernel/local.inc" --> 
<!DOCTYPE html>

<html>
  	<head>
    	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <title>Polla Mundial</title>
        <meta name="description" content="">
        <meta name="viewport" content="width=device-width, initial-scale=1">
		
        <style>
            body {
                background-color:#9F9F9F;
                background-image:url(Imagenes/fondo.jpg);
            }
			
            td {
                font-family:Verdana, Arial, Helvetica, sans-serif;
                font-size:15px;
            }
			
            #content h2 {
                text-align: left;
            }
        </style>
    </head>
	
    <body>
		<%
			dim k, campo		
		%>

		<table width="98%" style="margin-left: auto; margin-right: auto; padding: 0px; 
								  vertical-align: top;
								  border-style: none; border-spacing; 0px; background-color:#000000;">		
			<tr>
				<td>
					<img src="Imagenes/header.jpg" style="border-style: none; width: 100%;" />
				</td>
			</tr>

			<tr>
				<td style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 18px; text-align: center; color:rgb(255, 255, 255); padding: 10px;">
					Apuesta para la Polla Mundial
				</td>							
			</tr>

			<tr>
				<td style="padding: 0px; border-style: none; border-spacing: 0px; background-color: rgb(255, 255, 255);">
					<table style="width: 100%; margin-left: auto; margin-right: auto; padding: 0px; border-spacing: 0px; border-style: none;">
						<form name="DatosUsuario" action="grabar.asp" method="post">			

							<!-- 
								Aqui tenemos los 31 equipos que el usuario ha seleccionado en el mapa 
							-->

							<%
								for k = 1 to 31
									campo = "Polla" & k
									Response.Write "<input name='" & campo & "' type='hidden' value='" & request.Form(campo) & "'>" 										
								next
							%>
						
							<tr><td colspan="5">&nbsp;</td></tr>

							<!-- 
								Aqui van los campos del formulario... 
							-->
							<tr>
								<td width="15%">&nbsp;</td>
								<td width="25%" style="text-align:right;">Cedula:</td>
								<td width="2%">&nbsp;</td>
								<td width="34%" style="text-align:left;"><input id="txtCedula" name="txtCedula" type="text" size="15" maxlength="15"></td>
								<td width="24%">&nbsp;</td>							
							</tr>
							
							<tr><td colspan="5">&nbsp;</td></tr>
							
							<tr>
								<td width="15%">&nbsp;</td>
								<td width="25%" style="text-align:right;">Nombre:</td>
								<td width="2%">&nbsp;</td>
								<td width="34%" style="text-align:left;"><input id="txtNombre" name="txtNombre" type="text" size="25" maxlength="25"></td>
								<td width="24%">&nbsp;</td>							
							</tr>

							<tr><td colspan="5">&nbsp;</td></tr>
							
							<tr>
								<td width="15%">&nbsp;</td>
								<td width="25%" style="text-align:right;">Departamento:</td>
								<td width="2%">&nbsp;</td>
								<td width="34%" style="text-align:left;"><input id="txtDepartamento" name="txtDepartamento" type="text" size="25" maxlength="25"></td>
								<td width="24%">&nbsp;</td>							
							</tr>

							<tr><td colspan="5">&nbsp;</td></tr>
												
							<tr>
								<td width="15%">&nbsp;</td>
								<td width="25%" style="text-align:right;">Telefono / Ext: </td>
								<td width="2%">&nbsp;</td>
								<td width="34%" style="text-align:left;"><input id="txtTelefono" name="txtTelefono" type="text" size="15" maxlength="15"></td>
								<td width="24%">&nbsp;</td>							
							</tr>
							
							<tr>
								<td colspan="5">&nbsp;</td>
							</tr>

							<tr style="background-color:#E2E2E2; text-align:center;">
								<td colspan="5" style="padding: 15px;">
									<input name="btnSubmit" value=" Grabar Apuesta!" type="button" onClick="EnviarDatos()" style="font-size:18px;">
								</td>
							</tr>               
						</form>
					</table>
				</td>
			</tr>
		</table>
    </body>

	<script type="text/javascript">
		function EnviarDatos() {
			var sw = 0;
			
			if(document.getElementById("txtCedula").value.trim() == "") sw=1;
			if(document.getElementById("txtNombre").value.trim() == "") sw=1;
			if(document.getElementById("txtDepartamento").value.trim() == "") sw=1;
			if(document.getElementById("txtTelefono").value.trim() == "") sw=1;
			
			if(sw > 0) {
				alert("Uno o mas de los campos no ha sido llenado. Debe llenar todos los campos antes de grabar su apuesta!"); }
			else {
				document.DatosUsuario.submit();
			}
		}
	</script>	
</html>