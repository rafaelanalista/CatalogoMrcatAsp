	<%

Const adUseClient = 1
Const adOpenStatic = 3
Const adLockReadOnly = 1
Const adPersistXML = 1


	 set conn = Server.CreateObject("Adodb.Connection")
	 conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")
 	sql = "select produtos.referencia, produtos.ano from produtos "
   set rs = Server.CreateObject("adodb.recordset")	
rs.CursorLocation=adUseClient

rs.Open sql,conn,adOpenStatic,adLockReadOnly,adcmdtext

rs.Save  "c:\test.xml", adPersistXML
rs.close
set rs=nothing
conn.close
set conn=nothing
  


%>