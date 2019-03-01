<%

	 set conn = Server.CreateObject("Adodb.Connection")
	 conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")
       categoriaAtual = (request.queryString("categoria1"))

     sql ="select * from produtos inner join categorias on produtos.codcategoria=categorias.codcategoria where produtos.codcategoria=" & categoriaAtual  
      set rs = conn.execute(sql)
    
%>



<html>
	<head>
<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.8.1/jquery.min.js"></script>	
    
     <link href="//maxcdn.bootstrapcdn.com/font-awesome/4.2.0/css/font-awesome.min.css" rel="stylesheet">
	<script type="text/javascript" src="scrip.js"></script>





<style type=text/css>

#tudo{
width:980;
height:990;
margin: 0 auto;
border: 1px black solid;
}



#topo{

height:170;
}


#esquerda{
float:left;
width:270;
font-family: Arial;
height:800;
font-size:14px;
BORDER: 1PX BLACK SOLID;
}


#direito{
float:right;
width:680;

height:600;
BORDER: 1PX BLACK SOLID;

}
#busca{
background-color:grey;
text-align:right;
background-color:rgb(245,245,245)
}

#TOPO2{
FONT-FAMILY: lUCIDA;
text-align:right;

}

A{
text-decoration:none;
color:black;
}

a:hover {
background-color:black;
color:white;
}

input{
border-radius: 5px;	


}









</style>





	</head>
	<body>
         <div id=tudo>
			<div id=topo>
				<a href=index.asp><img src=logo.png></a>
			<br>
				<div id=busca>
				<form action = buscando.asp method=get>
					Referência <input type=text name=p>
					<input type=submit value=buscar onclick="submit();this.value='Carregando informações!'">

				</form>
			</div>
			<DIV ID=TOPO2>
				<B><A HREF=index.asp>WOMEN </A>| <a href=masculino.asp>MEN</a> | ACESSÓRIOS | BOLSAS</B>
			</DIV>
		 </div>
			<div id=esquerda>
				<!--#include file="menu.inc"-->
			</div>
			<div id=direito>
				jijijijijijiji		
				<!--#include file="menu2.inc"-->
<%	
	while not rs.eof
response.write rs.fields("FOTOP") 





rs.movenext
wend
%>				
			</div>
			
			
			
			
		 
		 
		 
		 
		 
		 </div>
      


	</body>





</html>	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	