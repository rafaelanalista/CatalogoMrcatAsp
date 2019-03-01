<%
	set conn = Server.CreateObject("Adodb.Connection")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")
 	sql = "select * from PRODUTOS inner join categorias on produtos.codcategoria=categorias.codcategoria WHERE produtos.codcategoria=11" 
    set tt = conn.execute(sql)
      nada = "nada encontrado"
%>
<html>
	<head>
		<meta charset="utf-8"> 
		<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.8.1/jquery.min.js"></script>	
        <link href="//maxcdn.bootstrapcdn.com/font-awesome/4.2.0/css/font-awesome.min.css" rel="stylesheet">
		<script type="text/javascript" src="scrip.js"></script>
		<link rel="stylesheet" style type="text/css"  href="estilo.css">
	</head>
	<body>
         <div id="tudo">
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
					IF tt.EOF THEN
						RESPONSE.WRITE "Ainda não foram cadastrados produtos desse modelo"
					ELSE
				%>
					<%=tt("CATEGORIA")%><BR>
				<%

					do while not tt.eof
				%>
					<a href="detalhes.asp?id=<%=tt.fields("id")%>"><img src="<%=tt("fotop")%>" width="230" height="200"></a>
  				<%
  					tt.movenext
  					loop
  					end if
  				%>
			</div>
		 </div>

<%
ID=100003308091638
PageLink="http://www.facebook.com/#!/profile.php?id==" & cstr(ID)

iFrameCode="<iframe src='http://www.facebook.com/plugins/like.php?href=" & PageLink & "&layout=standard&show-faces=true&width=160&action=like&font=tahoma&" &_
"colorscheme=light' scrolling='no' frameborder='0' allowTransparency='true' style='border:none; overflow:hidden; width:160px; height:90px'></iframe>"

response.write(iFrameCode)
%>

    </body>
</html>