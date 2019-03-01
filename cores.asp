<%
	set conn = Server.CreateObject("Adodb.Connection")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")

	sql = "select * from cor"
	set rs = conn.execute(sql)

	while not rs.eof
	%>
	
	<img src="<%=rs("foto")%>">

	<%
	

	rs.movenext
	wend




%>