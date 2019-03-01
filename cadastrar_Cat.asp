
<%
	set conn = Server.CreateObject("Adodb.Connection")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")

	sql = "select * from tipo"
	
	set rs = conn.execute(sql)

%>

<html lang='pt-BR'>
	<head>
		<link rel=stylesheet type=text/css href=estilo_admin.css>
		<title>Area Administrativa</title>
	</head>
	
	<body>
		<div id=conteudo>
			<header>
				<img src="logo.png">
			</header>
			<br />
			<h2>Área Administrativa</h2>
			<h3><a href="admin.asp">voltar ao menu</a></h3>
			<hr>
			<fieldset>
				<legend><h2>Cadastro de Categoria</h2></legend>
					<form action="cadastrando.asp" method="post">
						<br />Informe a Categoria: <input type="text" name=descricao required ><br />
						<br />Tipo de Produto: <select name="cat">
						<%
						do while not rs.eof
						%>
						<option value="<% = rs("codtipo")%>"><%=rs("tipo")%>
						</option>
						<%
						rs.movenext
						loop
						%>
						</select><br /><br />
						<input type=submit value=cadastrar>
					</form>
			</fieldset>
			
			
		</div>
	
	
	
	
	</body>
	



</html>