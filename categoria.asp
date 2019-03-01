<%
	set conn = Server.CreateObject("Adodb.Connection")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")
    categoriaAtual = (request.queryString("categoria1"))
    sqlquery ="select * from produtos inner join categorias on produtos.codcategoria=categorias.codcategoria where produtos.codcategoria=" & categoriaAtual  
    set executar = conn.execute(sqlquery)
%>
<html>
	<head>
		<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.8.1/jquery.min.js"></script>	
    	<link href="//maxcdn.bootstrapcdn.com/font-awesome/4.2.0/css/font-awesome.min.css" rel="stylesheet">
		<script type="text/javascript" src="scrip.js"></script>
		<link rel="stylesheet" style type="text/css"  href="estilo.css">
	</head>
	<body>
    	<div id=tudo>
			<div id="topo">
				<a href="index.asp"><IMG SRC="LOGO.png"></a>
				<div id="top">
					<form action="buscando.asp" method="get">
						<input type="text" name="p" class="busca" placeholder="Digite a referência ou modelo do produto e dâ um busca!">
						<input type="submit" value="buscar" class="btn" onclick="submit();this.value='Carregando informações!'">
					</form>
				</div>
				<div class="sexo">
					<a href="index.asp">Women</a> | <a href="masculino.asp">Men</a>
				</div>
			</div>
			<br>
			<div id="esquerda">
				<!--#include file="menu.inc"-->
			</div>
			<div id="direito">
				<I>Women's Collection</I><BR>	
				<HR>
				<%
					IF EXECUTAR.EOF THEN
						RESPONSE.WRITE "Ainda não foram cadastrados produtos nesse modelo"
					ELSE
				%>
					<%=EXECUTAR("CATEGORIA")%><BR>





				<% while not executar.eof %>
					<a href="detalhes.asp?id=<%=executar.fields("id")%>"><img src="<%=executar("fotop")%>" width="230" height="200"></a>				
				<%
					executar.movenext
					wend
					end if
				%>
			</div>
		</div>
	</body>
</html>	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	