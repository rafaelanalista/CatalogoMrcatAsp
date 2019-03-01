
<%
	set conn = Server.CreateObject("Adodb.Connection")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")
sql = "select tipo.tipo, categorias.categoria, codcategoria from tipo inner join categorias on tipo.codtipo =  categorias.codtipo"
set rs = conn.execute(sql)


%>

<html lang='pt-BR'>
	<head>
		<link rel=stylesheet type=text/css href=estilo_admin.css>
		<title>Area Administrativa</title>
			


		<SCRIPT>
			function confirm_delete(){
				return confirm ("Tem certeza que deseja excluir?")
			}
		</script>

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
			<table border=1>
				<tr ><td bgcolor=black color=white><h4>categoria</h4></td>
					<td bgcolor=black color=white><h4>tipo</h4></td>
					<td bgcolor=black color=white><h4>excluir</h4></td>
				</tr>
				
					<% while not rs.eof%>
			
				<tr>
					<td><%=rs.fields("categoria")%></td>
					<td><%=rs.fields("tipo")%></td>
					<td><a href="excluindo.asp?codcategoria1=<%=rs.fields("codcategoria")%>"><img src=btexcluir.PNG  WIDTH=50 HEIGHT=50 onclick="return confirm_delete()"></a> </td>
				</tr> 
			
			<%
				rs.movenext
			wend

				

			%>
			
			</table>
		</div>
	
	
	
	
	</body>
	



</html>