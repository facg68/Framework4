<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    Snip_Width = 500    
%>
    
<!-- #include virtual = "/core/includes/snippets.inc" -->

<%
    dim gfx_ga_snip_con, gfx_ga_snip_t, gfx_ga_snip_sqlString
    dim gfx_ga_snip_data, gfx_ga_snip_labels

    set gfx_ga_snip_con = Server.CreateObject("ADODB.Connection")
    gfx_ga_snip_con.open Application("Conn")

    Usuario = Request.Cookies("Usuario")
    gfx_ga_snip_Amo = YEAR(NOW())     

    sqlString = "exec discos_gfx_gastos '" & Usuario & "', 1, 2"       
    apexArea "", sqlString, "Año", "Total", "#008FFB", 250, Snip_Width
  
    gfx_ga_snip_con.close: set gfx_ga_snip_con = nothing 
%> 

<script>
    function gfx_gastos_anuales_init() {
        redimWindow("gfx_gastos_anuales", <%= Snip_Width %>)
    }
</script>     