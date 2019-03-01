
<%
	set conn = Server.CreateObject("Adodb.Connection")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")


%>

<html lang='pt-BR'>
	<head>
		<link rel=stylesheet type=text/css href=estilo_admin.css>
		<title>Area Administrativa</title>
	</head>
	
	<body>
		<div id=conteudo>
			<header>
				<img src=logo.png>
			</header>
			<br />
			<h2>Área Administrativa</h2>
			<a href=admin.asp><h3>voltar ao menu</h3></a>
			<hr>
			<%
			codexclui = Request.QueryString("id1")
			sql = "delete from produtos where id=" &codexclui
			set rs = conn.execute(sql)


			%>

			<script>
				alert("produto excluido com sucesso");
			</script>
			<%
				response.write "<meta http-equiv='refresh' content='0; url=excluir_produtos.asp'/>"
			%>


			
			
			
		</div>
	
	
	
	
	</body>
	



</html>