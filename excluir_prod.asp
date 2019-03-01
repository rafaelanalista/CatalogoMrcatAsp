<%
varreferencia = Request.form("txtreferencia")
set conn = Server.CreateObject("Adodb.Connection")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" &Server.MapPath("mrcat.mdb")
set rs =Server.CreateObject("Adodb.Recordset")

	rs.Open "Select categorias.categoria, produtos.referencia, produtos.id, produtos.cor, produtos.ano, produtos.fotop, produtos.fotog from categorias inner join produtos on categorias.codcategoria=produtos.codcategoria where produtos.referencia like '%"&varreferencia&"%'",conn,3,3

	rs.PageSize = 4


%>
<html lang='pt-BR'>
   <head>
      <meta charset = utf-8>
      <link rel=stylesheet type=text/css href=estilo_admin.css>
      <title>Area Administrativa</title>
      <style type="text/css">
      body{
         font-family:arial;
      }

      </style>
            <SCRIPT>
         function confirm_delete(){
            return confirm ("Tem certeza que deseja excluir?")
         }
      </script>



   </head>
   
   <body>
      <div id=conteudo>
         <header>
            <img src=logo.png>
         </header>
         <br />
         <h2>Área Administrativa</h2>
         <a href=excluir_Produtos.asp><h3>voltar ao menu</h3></a>
         <hr>



<table  >
<tr height="40">
<td width="195" bgcolor=black ><font color=white>Referencia</font></td>

<td width="100" bgcolor=black ><font color=white>Cor</td>
<td width="110" bgcolor=black ><font color=white>Ano</td>
<td width="80" bgcolor=black ><font color=white>Categoria</td>
<td width="210" bgcolor=black ><font color=white>Foto </td>

</tr>
</table>


<table border=2>
<%
IF RS.EOF then 
   Response.Write "nenhum registro encontrado"
   Response.End

ELSE
   IF Request.QueryString("pagina")=" " then 
      intpagina=1 
   ELSE
      IF cint(Request.QueryString("pagina"))<1 then
         intpagina=1

      ELSE
         IF cint(Request.QueryString("pagina"))> RS.PageCount then 
            intpagina=RS.PageCount

         ELSE
            intpagina=Request.QueryString("pagina")
         END IF
      END IF
   END IF
END IF
rs.AbsolutePage=intpagina
intrec=0
While intrec < rs.PageSize and not RS.EOF 
%>
<tr> 
<td width=190><%=rs("referencia")%></td>
<td width=100><%=rs("cor")%></td>
<td width=92><%=rs("ano")%></td>
<td width=92><%=rs("categoria")%></td>
<td width=100><img width=200 height=200 src="<%=rs("fotop")%>"></td>
<td width=92><a href="excluindoprod.asp?id1=<%=rs.fields("id")%>"><img src=btexcluir.jpg onclick="return confirm_delete()"></a></td>


<td>
</tr>


<%
rs.movenext
ntrec=intrec+1 
   IF RS.EOF then 
      response.write " "

   END IF
wend
IF intpagina > 1 then 
%>

<a href="excluir_prod.asp?pagina=<%=intpagina-1%>"><b>/ Anterior</b></a>
<% 
END IF
IF strcomp(intpagina,RS.PageCount) <> 0 then 
%> 
   <a href="excluir_prod.asp?pagina=<%=intpagina+ 1%>"><b>Próxima</b></a> 
<% 
END IF
%>
</table>
</div>
</html>
</html>