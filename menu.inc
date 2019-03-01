	<%
    	set conn = Server.CreateObject("Adodb.Connection")
	    conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")
   		sqlbolsas = "select * from categorias where codtipo = 5 "
	    set queri = conn.execute(sqlbolsas)
        sqlsapatosmasc = "select * from categorias where codtipo = 6 "
        set sapatomasc = conn.execute(sqlsapatosmasc)
        sqlacessfem = "select * from categorias where codtipo = 4 "
        set acessfem = conn.execute(sqlacessfem)

        query = "select * from cor"
	    set cor = conn.execute(query)

	%>
    <div id="accordion">
		<h4 class="accordion-toggle">SAPATOS</h4>
	    <div class="accordion-content default">
			<%




  				while not sapatomasc.eof
			%>
  			&nbsp&nbsp<a href="categoria.asp?categoria1=<%=sapatomasc.fields("codcategoria")%>"><%=sapatomasc.fields("categoria")%></a> <br><br>   
			<%
				sapatomasc.movenext
				wend
			

			%>
		</div>
		<h4 class="accordion-toggle">BOLSAS</h4>
		<div class="accordion-content ">
			<%
  			while not queri.eof
  			%>
			&nbsp&nbsp<a href="categoria.asp?categoria1=<%=queri.fields("codcategoria")%>"><%=queri.fields("categoria")%></a> <br><br>
			<%
   			queri.movenext
			wend
			%>
		</div>
		<h4 class="accordion-toggle">ACESSÓRIOS</h4>
		<div class="accordion-content ">
			<%
  			while not acessfem.eof
  			%>
			&nbsp&nbsp<a href="categoria.asp?categoria1=<%=acessfem.fields("codcategoria")%>"><%=acessfem.fields("categoria")%></a> <br><br>
			<%
   			acessfem.movenext
			wend
			%>
		</div>


				<h4 class="accordion-toggle">CORES</h4>
		<div class="accordion-content ">

			<%
				while not cor.eof

			%>
					<img src="<%=cor("foto")%>">

			<%
				cor.movenext
				wend


			%>
		</div>







	</div>
  
	  
  