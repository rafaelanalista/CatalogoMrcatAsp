<head>
	<head>
		<meta charset = "utf-8">
    	<title>Pagina&ccedil;&atilde;o</title>
		<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.8.1/jquery.min.js"></script>	
		<link href="//maxcdn.bootstrapcdn.com/font-awesome/4.2.0/css/font-awesome.min.css" rel="stylesheet">
		<script type="text/javascript" src="scrip.js"></script>
		<link rel="stylesheet" style type="text/css"  href="estilo.css">
	</head>
	<body >

		<div id=tudo>
			<div id="topo">
				<a href=index.asp><IMG SRC=LOGO.png  ></a>
				<div id=top>
					<form action=buscando.asp method=get>
						<input type=text name="p" class=busca placeholder="Digite a referência ou modelo do produto e dâ um busca!">
						<input type=submit value=buscar class=btn onclick="submit();this.value='Carregando informações!'">
					</form>
				</div>
				<div class="sexo">
					<a href=index.asp>Women</a> | <a href=masculino.asp>Men</a>
				</div>
			</div>

			
			<br>
			<div id="esquerda">
				<!--#include file="menumasc.inc"-->	
			</div>
  			<div id="direita">
  				<hr />
				<% 
					ON ERROR RESUME NEXT
					Dim sql
					Dim ul
					Dim vExibe
					Dim txt
					Dim vCrt
					Dim StrIngredientes
					Dim busca
					Dim intpagina
					Dim intrec
					Dim i
					set DB = Server.CreateObject("Adodb.Connection")
					DB.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")
					busca=request("p")
					'objeto recordset
					Set RS = Server.CreateObject("Adodb.RecordSet")
					RS.PageSize = 9
					sql = "SELECT * FROM produtos where referencia LIKE '%"&busca&"%'"
					RS.Open sql, DB,3,3
					If RS.EOF then
						Response.Write "Nenhum registro encontrado!"
						Response.End
					Else
					If Request.QueryString("pagina")="" then
						intpagina=1
					Else
					If Cint(Request.QueryString("pagina"))<1 then
						intpagina=1
					Else
					If Cint(Request.QueryString("pagina"))>RS.PageCount then
						intpagina=RS.PageCount
					Else
					intpagina=Request.QueryString("pagina")
					End If
					End if
					End if
					End if 
					RS.AbsolutePage=intpagina
					intrec=0
				%>
				Palavra pesquisada: <b><%=busca%></b><br />
				Foram encontrados  <b> <%=rs.recordcount%> </b>registros <br /><br />
				Pagina <%=intpagina%> de <%=RS.PageCount%></p>
				
				</b> 
				<div align="center">
					<center>
						<%
							While intrec<RS.PageSize and not RS.EOF
						%>
						<a href="detalhes.asp?id=<%=rs.fields("id")%>"><img height=220 width=220 src="<%=RS("fotop")%>"></a>
						<%
							Rs.movenext
							intrec=intrec+1
							Wend
						%>
					</center>
				</div>
				<br>
				<div id="count" align="center">
					<%
						If intpagina > 1 then
							Response.Write "<a href=""buscando.asp?pagina=" & intpagina-1 & """>Anterior</a>"
						End If
						For i=1 to RS.PageCount
						If i = Cint(intpagina) then
							Response.Write " " & i
						else
							Response.Write " <a href=""buscando.asp?pagina=" & i &"&p="&busca& """>" & i & "</a>"
						End If
						Next
						If strcomp(intpagina,RS.PageCount)<> 0 then
							Response.Write " <a href=""buscando.asp?pagina=" & intpagina+1 & """>Próxima</a>"
						End If
						DB.close
					%>
			</div>
		</div>
	</body>
</html>