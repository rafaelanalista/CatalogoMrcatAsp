<%
	set DB = Server.CreateObject("Adodb.Connection")
	DB.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")
	id = request("id")
 	sql = "select * from produtos where id="&id
	set rs = DB.execute(sql)
%>
<html lang='pt-BR'>
<head>
	<meta charset = "utf-8">		
    <title>Pagina&ccedil;&atilde;o</title>
	<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.8.1/jquery.min.js"></script>	
	<link href="//maxcdn.bootstrapcdn.com/font-awesome/4.2.0/css/font-awesome.min.css" rel="stylesheet">
    <script type="text/javascript" src="scrip.js"></script>
    <script src='js/jquery-1.8.3.min.js'></script>
	<script src='js/jquery.elevatezoom.js'></script>
	<link rel="stylesheet" style type="text/css"  href="estilo.css">
	<style type="text/css">
		#direito{
			font-family:arial;


		}

	</style>


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
			<%
				while not rs.eof
			%>
			<img borde="1" src="<%=rs("fotop")%>" width="240" height="210" id="zoom_01" class="elevate-image" data-zoom-image="<%=rs("fotog")%>">
			<br>
			<br>
			<br>
			Cor: <b><%=rs("cor")%></b><br><br>
			Referência: <b><%=rs("referencia")%></b>
			<script>
    			$('#zoom_01').elevateZoom(); 
			</script>
			<br><br>
			Descrição do produto:
			<b>
			<%
				if rs("descricao") <> "" then
			%>
				<%=rs("descricao")%>
			<%
				else
					response.write "sem descrição"
				end if
			%>
			<%
 				rs.movenext
				wend
			%>
			</B>
		</div>
	</div>
</body>
</html>