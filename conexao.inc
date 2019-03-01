<%
	sub AbreConexao
		set conn = Server.CreateObject("Adodb.Connection")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")


	end sub
	
	sub FechaConexao
		conn.close
		set conn = nothing
	end sub



%>