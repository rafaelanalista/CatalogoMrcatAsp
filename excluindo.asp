
<%
	on Error Resume Next


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
			codexclui = Request.QueryString("codcategoria1")
			sql = "delete from categorias where codcategoria=" &codexclui
			set rs = conn.execute(sql)

			if Err.Number <> 0 then
				response.write "Houve um erro"

			
			%>

			<%
			else
			%>

			<script>
				alert("Categoria excluida com sucesso");
			</script>
			<%
			 end if 
			%>


			<%
				response.write "<meta http-equiv='refresh' content='0; url=excluir_Cat.asp'/>"




			%>


			
			
			
		</div>
	
	
	
	
	</body>
	



</html>