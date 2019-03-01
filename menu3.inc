<%

	 set conn = Server.CreateObject("Adodb.Connection")
	 conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")
       categoriaAtual=(request.queryString("categoria1"))

     sql ="select * from produtos inner join categorias on categorias.codcategoria=produtos.codcategoria where produtos.codcategoria=" &categoriaAtual  
      set rs = conn.execute(sql)
      while not rs.eof
	  %>

<div>
<img src ="<%=RS("fotop")%>">
</div>
<%

rs.movenext
wend
%>