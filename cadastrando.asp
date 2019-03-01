<%

set conn = Server.CreateObject("Adodb.Connection")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")

categoria = request.form("descricao")
codigo = request.form("cat")

sql = "Insert into categorias (categoria,codtipo) values ('"& categoria &"', '"& codigo &"')"
set rs = conn.execute(sql)


%>
 
<script>
alert("produto cadastrado com sucesso");
</script>
<%
response.write "<meta http-equiv='refresh' content='0; url=cadastrar_Cat.asp'/>"
%>
