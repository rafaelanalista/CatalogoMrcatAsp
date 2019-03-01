<%
	set conn = Server.CreateObject("Adodb.Connection")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")
 	sql = "select * from PRODUTOS inner join categorias on produtos.codcategoria=categorias.codcategoria WHERE produtos.codcategoria=61"
    set tt = conn.execute(sql)
%>
<html>
	<head>
		<meta charset = "utf-8"> 
		<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.8.1/jquery.min.js"></script>	
    	<link href="//maxcdn.bootstrapcdn.com/font-awesome/4.2.0/css/font-awesome.min.css" rel="stylesheet">
		<script type="text/javascript" src="scrip.js"></script>
		<link rel="stylesheet" style type="text/css"  href="estilo.css">
	</head>
	<body>
    	<div id="tudo">
			<div id="topo">
				<a href="index.asp"><IMG SRC=LOGO.png></a>
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
				<!--#include file="menumasc.inc"-->
			</div>
			<div id="direito">
				<I>Men's Collection</I><BR>
				<hr>
				<%
					IF tt.EOF THEN
						RESPONSE.WRITE "Ainda não foram cadastrados produtos desse modelo"
					ELSE
				%>
					<%=tt("CATEGORIA")%><BR>


				<%
				while not tt.eof
				%>
				<a href="detalhes.asp?id=<%=tt.fields("id")%>"><img src="<%=tt("fotop")%>" width="230" height="200"></a>
  			    <%
  				tt.movenext
  				wend
  				end if
  				%>
			</div>
		</div>
    </body>
</html>