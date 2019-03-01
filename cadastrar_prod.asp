<%
	set conn = Server.CreateObject("Adodb.Connection")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")
	sql = "select * from categorias"
	set rs = conn.execute(sql)





%>
<html lang='pt-BR'>
	<head>
		<meta charset = UTF-8>
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
			<fieldset>
				<legend><h2>Cadastro de Produto</h2></legend>
						<table>
							<form action=cadastrando_prod.asp method=post enctype="multipart/form-data">
							<tr>
							<td>Informe a refêrencia do produto:</td>
							<td><input type=text name=ref required ></td>
							</tr>
							<tr>
							<td>Modelo do Produto:</td>
                            <td><select name=txtcat>
						<%
						do while not rs.eof
						%>
						<option value="<% = rs("codcategoria")%>"><%=rs("categoria")%>
						</option>
						<%
						rs.movenext
						loop
						%>
						</select></td>
						</tr>
						<tr>
						<td>Informe a cor do produto:</td>
						<td><input type=text name=cor required ></td>
						</tr>
						<td>Informe o ano do produto:</td>
						<td><input type=text name=ano required ></td>
						</tr>
						<tr>
						<td>Foto pequena:</td>
						<td><input type=file name=foto size=14></td>
						</tr>
						<tr>
						<td>Foto grande:</td>
						<td><input type=file name=fotog size=14></td>
						</tr>					
						<tr>
						<td>Descriçao</td>
						<td><textarea rows="10" cols="40" maxlength="500"></textarea></td>
					</tr>
						<tr>
						<td><input type=submit value=cadastrar></td>
						</tr>



					</form>
					</table>
			</fieldset>
			
			
		</div>
	
	
	
	
	</body>
	



</html>