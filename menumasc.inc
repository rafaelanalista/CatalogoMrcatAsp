	<%
    	set conn = Server.CreateObject("Adodb.Connection")
	    conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")
   		sqlsapmasc = "select * from categorias where codtipo = 3 "
	    set sapmasc = conn.execute(sqlsapmasc)
        sqlacess = "select * from categorias where codtipo = 1 "
        set acess = conn.execute(sqlacess)
        
	%>
    <div id="accordion">
		<h4 class="accordion-toggle">SAPATOS</h4>
	    <div class="accordion-content default">
			<%
  				while not sapmasc.eof
			%>
  			&nbsp&nbsp<a href="categoriamasc.asp?categoria2=<%=sapmasc.fields("codcategoria")%>"><%=sapmasc.fields("categoria")%></a> <br><br>   
			<%
				sapmasc.movenext
				wend
			%>
		</div>
		<h4 class="accordion-toggle">ACESSÓRIOS</h4>
		<div class="accordion-content ">
			<%
  			while not acess.eof
  			%>
			&nbsp&nbsp<a href="categoriamasc.asp?categoria2=<%=acess.fields("codcategoria")%>"><%=acess.fields("categoria")%></a> <br><br>
			<%
   			acess.movenext
			wend
			%>
		</div>
		
	</div>
  
  
  