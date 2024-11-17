<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Dim adoCnn
Dim rsCadastro
Dim sCnn, sSQL

Tipo = Request.QueryString("Tipo")
Email = Request.QueryString("Email")
Bairro = Request.Form("txtBairro")
Cidade = Request.Form("txtCidade")
Uf = Request.Form("cboUf")

Set adoCnn = Server.CreateObject("ADODB.Connection")
Set rsCadastro = Server.CreateObject("ADODB.RecordSet")
sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../db/dbGG.mdb")
'--sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\web\localuser\gymgroup\banco\dbGG.mdb"--'
adoCnn.Open(sCnn)

sSQL = "SELECT * FROM Cadastro WHERE cadastroEmail='" & Email & "'"
rsCadastro.Open sSQL, adoCnn, 1, 1
Nome = rsCadastro("cadastroNome")
rsCadastro.Close

IF Tipo = "PROFISSIONAL" THEN
	IF Uf = EMPTY AND Cidade = EMPTY AND Bairro = EMPTY THEN
		sSQL = "SELECT * FROM Cadastro WHERE cadastroTipo<>'PROFISSIONAL' ORDER BY cadastroNome, cadastroUf, cadastroCidade, cadastroBairro"
	ELSEIF NOT ISEMPTY(Uf) AND Cidade = Empty AND Bairro = Empty THEN
		sSQL = "SELECT * FROM Cadastro WHERE cadastroTipo<>'PROFISSIONAL' AND cadastroUf='" & Uf & "' ORDER BY cadastroNome, cadastroUf, cadastroCidade, cadastroBairro" 
	ELSEIF NOT ISEMPTY(Uf) AND NOT ISEMPTY(Cidade) AND Bairro = Empty THEN
		sSQL = "SELECT * FROM Cadastro WHERE cadastroTipo<>'PROFISSIONAL' AND cadastroUf='" & Uf & "' AND cadastroCidade='" & Cidade & "' ORDER BY cadastroNome, cadastroUf, cadastroCidade, cadastroBairro"
	ELSEIF NOT ISEMPTY(Uf) AND NOT ISEMPTY(Cidade) AND NOT ISEMPTY(Bairro) THEN
		sSQL = "SELECT * FROM Cadastro WHERE cadastroTipo<>'PROFISSIONAL' AND cadastroUf='" & Uf & "' AND cadastroCidade='" & Cidade & "' AND cadastroBairro='" & Bairro & "' ORDER BY cadastroNome, cadastroUf, cadastroCidade, cadastroBairro"
	END IF
ELSE
	IF Uf = EMPTY AND Cidade = EMPTY AND Bairro = EMPTY THEN
		sSQL = "SELECT * FROM Cadastro WHERE cadastroTipo='PROFISSIONAL' ORDER BY cadastroNome, cadastroUf, cadastroCidade, cadastroBairro"
	ELSEIF NOT ISEMPTY(Uf) AND Cidade = Empty AND Bairro = Empty THEN
		sSQL = "SELECT * FROM Cadastro WHERE cadastroTipo='PROFISSIONAL' AND cadastroUf='" & Uf & "' ORDER BY cadastroNome, cadastroUf, cadastroCidade, cadastroBairro" 
	ELSEIF NOT ISEMPTY(Uf) AND NOT ISEMPTY(Cidade) AND Bairro = Empty THEN
		sSQL = "SELECT * FROM Cadastro WHERE cadastroTipo='PROFISSIONAL' AND cadastroUf='" & Uf & "' AND cadastroCidade='" & Cidade & "' ORDER BY cadastroNome, cadastroUf, cadastroCidade, cadastroBairro"
	ELSEIF NOT ISEMPTY(Uf) AND NOT ISEMPTY(Cidade) AND NOT ISEMPTY(Bairro) THEN
		sSQL = "SELECT * FROM Cadastro WHERE cadastroTipo='PROFISSIONAL' AND cadastroUf='" & Uf & "' AND cadastroCidade='" & Cidade & "' AND cadastroBairro='" & Bairro & "' ORDER BY cadastroNome, cadastroUf, cadastroCidade, cadastroBairro"
	END IF
END IF

rsCadastro.Open sSQL, adoCnn, 1, 1
%>
<!doctype html>

<html lang="pt-br">

<head>
<meta charset="UTF-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<meta name="description" content="Seu Grupo Saudável">
	<meta name="keywords" content="Grupo, Saúde, Exercício, Corrida, Musculação, Fitness, Ginástica, Caminhada, Crossfit, Físico, Academia">
	<link rel="stylesheet" href="../css/gymgroup.css">
	<title>GYM GROUP :: Pesquisar</title>
</head>
	
<body leftmargin="5" topmargin="5">
<table width="800" align="center">
  <tr>
		<td width="680"><img src="../imgs/logo.png" width="118" height="90" alt=""/></td>
	  <td width="70" align="right"><a href="javascript:history.back()">VOLTAR</a></td>
	  <td width="50" align="right"><a href="../index.asp">SAIR</a></td>
  </tr>
</table>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr height="20" class="subttl">
	<td width="75%">Usuário:&nbsp;<span class="selo"><%=Response.Write(Nome)%></span></td>
	<td width="25%" align="right">Tipo:&nbsp;<span class="selo"><%=Response.Write(Tipo)%></span></td>
  </tr>
</table>
<br>
<%IF rsCadastro.RecordCount = 0 THEN%>
<div align="center"><p class="aviso">Não Existem Cadastros com estes Parâmetros!<br>
Realize uma Nova Pesquisa!<br>
<br></p></div>
  <%rsCadastro.Close%>
  <%SET rsCadastro = Nothing%>
  <%adoCnn.Close%>
  <%SET adoCnn = Nothing%>
<%ELSE%>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr class="subttl">
    <td width="40%">Nome</td>
    <td width="25%">Bairro</td>
	<td width="25%">Cidade</td> 
    <td width="10%">UF</td>
  </tr>
  <%rsCadastro.MoveFirst%>
  <%FOR I = 1 TO rsCadastro.RecordCount%>
  <tr>
  	<td valign="middle"><p class="aviso"><a href="dados.asp?sEmail01=<%=Email%>&sEmail02=<%=rsCadastro("cadastroEmail")%>">+</a><%=Response.Write(rsCadastro("cadastroNome"))%></p></td>
  	<td valign="middle"><p class="aviso"><%=Response.Write(rsCadastro("cadastroBairro"))%></p></td>
	<td valign="middle"><p class="aviso"><%=Response.Write(rsCadastro("cadastroCidade"))%></p></td>
  	<td valign="middle"><p class="aviso"><%=Response.Write(rsCadastro("cadastroUf"))%></p></td>
  </tr>
  <%rsCadastro.MoveNext%>
  <%NEXT%>
  <%rsCadastro.Close%>
  <%SET rsCadastro = Nothing%>
  <%adoCnn.Close%>
  <%SET adoCnn = Nothing%>
</table>
<%END IF%>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<div class="aviso" align="center">GYM GROUP <%Response.Write("2024" & "-" & Year(Now))%> © Todos os Direitos Reservados</div>
</body>
</html>
