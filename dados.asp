<%@LANGUAGE="VBSCRIPT"%>
<%
Dim adoCnn
Dim rsCadastro
Dim sCnn, sSQL
Dim Nome
Dim Tipo
Dim Email01, Email02

Email01 = Request.QueryString("sEmail01")
Email02 = Request.QueryString("sEmail02")

Set adoCnn = Server.CreateObject("ADODB.Connection")
Set rsCadastro = Server.CreateObject("ADODB.RecordSet")
sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../db/dbGG.mdb")
adoCnn.Open(sCnn)

sSQL = "SELECT * FROM Cadastro WHERE cadastroEmail='" & Email01 & "'"
rsCadastro.Open sSQL, adoCnn, 1, 1
Nome = rsCadastro("cadastroNome")
Tipo = rsCadastro("cadastroTipo")
rsCadastro.Close

sSQL = "SELECT * FROM Cadastro WHERE cadastroEmail='" & Email02 & "'"
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
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="rotulo" width="10%"><p>Nome</p></td>
	<td class="aviso" width="90%"><p><%=Response.Write(rsCadastro("cadastroNome"))%></p></td>
  </tr>
	<tr>
    <td class="rotulo" width="10%"><p>Telefone</p></td>
	<td class="aviso" width="90%"><p><%=Response.Write(rsCadastro("cadastroCelular"))%></p></td>
  </tr>
	<tr>
    <td class="rotulo" width="10%"><p>e-Mail</p></td>
	<td class="aviso" width="90%"><p><%=Response.Write(rsCadastro("cadastroEmail"))%></p></td>
  </tr>
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
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<div class="aviso" align="center">Todos os Direitos Reservados <%Response.Write("2024" & "-" & Year(Now))%> © GYM GROUP</div>
</body>
</html>
