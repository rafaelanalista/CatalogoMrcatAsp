<!-- #include file = "upload_funcoes.asp" -->
<%
' Chamando Funções, que fazem o Upload funcionar
byteCount = Request.TotalBytes
RequestBin = Request.BinaryRead(byteCount)
Set UploadRequest = CreateObject("Scripting.Dictionary")
BuildUploadRequest RequestBin

' Recuperando os Dados Digitados ----------------------
referenci = UploadRequest.Item("ref").Item("Value")
categ = UploadRequest.Item("txtcat").Item("Value")
co = UploadRequest.Item("cor").Item("Value")
an = UploadRequest.Item("ano").Item("Value")



' Tipo de arquivo que esta sendo enviado
tipo_foto = UploadRequest.Item("foto").Item("ContentType")

' Caminho completo dos arquivos enviados
caminho_foto = UploadRequest.Item("foto").Item("FileName")

' Nome dos arquivos enviados
nome_foto = Right(caminho_foto,Len(caminho_foto)-InstrRev(caminho_foto,"\"))

' Conteudo binario dos arquivos enviados
foto = UploadRequest.Item("foto").Item("Value")

' pasta onde as imagens serao guardadas
pasta = Server.MapPath("images/")
nome_foto = "/"&nome_foto

' pasta + nome dos arquivos
cfoto = "images" + nome_foto

' Fazendo o Upload do arquivo selecionado
if foto <> "" then
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
Set MyFile = ScriptObject.CreateTextFile(pasta & nome_foto)
For i = 1 to LenB(foto)
MyFile.Write chr(AscB(MidB(foto,i,1)))
Next
MyFile.Close
end if











' Tipo de arquivo que esta sendo enviado
tipo_fot = UploadRequest.Item("fotog").Item("ContentType")

' Caminho completo dos arquivos enviados
caminho_fot = UploadRequest.Item("fotog").Item("FileName")

' Nome dos arquivos enviados
nome_fot = Right(caminho_fot,Len(caminho_fot)-InstrRev(caminho_fot,"\"))

' Conteudo binario dos arquivos enviados
fotog = UploadRequest.Item("fotog").Item("Value")

' pasta onde as imagens serao guardadas
pasta = Server.MapPath("images/")
nome_fot = "/"&nome_fot

' pasta + nome dos arquivos
dfoto = "images" + nome_fot

' Fazendo o Upload do arquivo selecionado
if fotog <> "" then
Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
Set MyFile = ScriptObject.CreateTextFile(pasta & nome_fot)
For i = 1 to LenB(fotog)
MyFile.Write chr(AscB(MidB(fotog,i,1)))
Next
MyFile.Close
end if






' Conecta-se ao Banco de Dados
url_conexao = Server.MapPath("mrcat.mdb")
set conexao = Server.CreateObject("ADODB.Connection")
conexao.open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ="&url_conexao	

' cadastra os dados no banco de dados
sql = "Insert into produtos (referencia,codcategoria,cor,ano,fotop, fotog) values ('"& referenci &"', '"& categ &"', '"& co &"', '"& an &"', '"& cfoto &"', '"& dfoto &"')"
Conexao.Execute(sql)

' Mostra Mensagem de Confirmação na Tela
Response.write "Dados Cadastrados com Sucesso!"

' Redireciona após 5 segundos
response.write "<br><br>você será redirecionado em 5 segundos..<br>"
response.write "<meta http-equiv='refresh' content='0; url=admin.asp'/>"
%>
