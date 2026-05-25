<%
    Response.CodePage = 65001
    Response.Charset = "UTF-8"

    Snip_Width = 500       
%>

<!-- #include virtual = "/core/includes/snippets.inc" -->

<%
    dim gfx_ma_snip_con, gfx_ma_snip_t, gfx_ma_snip_sqlString
    dim gfx_ma_snip_data, gfx_ma_snip_labels

    set gfx_ma_snip_con = Server.CreateObject("ADODB.Connection")
    gfx_ma_snip_con.open Application("Conn")

    Usuario = Request.Cookies("Usuario")
    gfx_ma_snip_Amo = YEAR(NOW())

    sqlString = "exec discos_gfx_anuales '" & Usuario & "', 1, 2"      
    apexColumns "", "chart", sqlString, "Año", "Cantidad", "#064f8aff", 250, Snip_Width     
               
    gfx_ma_snip_con.close: set gfx_ma_snip_con = nothing 
%> 

<script>
    function gfx_medios_anuales_init() {
        redimWindow("gfx_medios_anuales", <%= Snip_Width %>)
    }
</script>      