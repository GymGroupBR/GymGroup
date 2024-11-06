<%@LANGUAGE="VBSCRIPT"%>
<%
Dim adoCnn
Dim rsCadastro
Dim sCnn, sSQL

Nome = Request.QueryString("Nome")
Senha = Request.QueryString("Senha")
Tipo = Request.QueryString("Tipo")
Email = Request.QueryString("Email")
Bairro = Request.Form("txtBairro")
Cidade = Request.Form("txtCidade")
Uf = Request.Form("cboUf")

Set adoCnn = Server.CreateObject("ADODB.Connection")
Set rsCadastro = Server.CreateObject("ADODB.RecordSet")
sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../db/dbGG.mdb")
adoCnn.Open(sCnn)

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
	<title>GYM GROUP ::: Pesquisar</title>
</head>
	
<body leftmargin="5" topmargin="5">
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
		<td align="center"><img src="../imgs/logo.png" width="235" height="179" alt=""/></td>
	</tr>
	</table>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr height="20">
    <td width="90%" bgcolor="#4260AC"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFCC00"><b>&nbsp;Usuário: <font color="#FFFFFF"><%=Response.Write(Nome)%> (<%=Response.Write(Tipo)%>)</font></b></font></td>
	<td width="10%" bgcolor="#4260AC" align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><a href="../index.asp" onMouseOver="MM_displayStatusMsg('..:: GYM GROUP :: Sair ::..');return document.MM_returnValue">SAIR</a>&nbsp;</b></font></td>
  </tr>
</table>
<br>
<%IF rsCadastro.RecordCount = 0 THEN%>
<div align="center"><p class="aviso">Não Existem Cadastros com estes Parâmetros!<br>
Realize uma Nova Pesquisa!<br>
<br>
<a class="aviso" href="javascript:history.back()" onMouseOver="MM_displayStatusMsg('..:: GYM GROUP :: Pesquisa ::..');return document.MM_returnValue">Voltar</a></p></div>
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
<br>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="aviso" align="center"><a href="javascript:history.back()" onMouseOver="MM_displayStatusMsg('..:: GYM GROUP :: Pesquisar ::..');return document.MM_returnValue">Voltar</a></td>
  </tr>
</table>
<%END IF%>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<div class="aviso" align="center">Todos os Direitos Reservados <%Response.Write("2024" & "-" & Year(Now))%> © GYM GROUP</div>
</body>
</html>
