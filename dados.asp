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
'--sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../db/dbGG.mdb")--'
sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "D:\web\localuser\gymgroup\banco\dbGG.mdb"
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
<table width="800" align="center">
  <tr>
		<td width="560"><img src="../imgs/logo.png" width="118" height="90" alt=""/></td>
	  <td width="120" align="right"><a href="cadastroIndividual.asp?sEmail=<%=Email%>">MEU CADASTRO</a></td>
	  <td width="70" align="right"><a href="javascript:history.back()">VOLTAR</a></td>
	  <td width="50" align="right"><a href="../index.asp">SAIR</a></td>
  </tr>
</table>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<table width="800" align="center" border="0" cellspacing="0" cellpadding="0"></span>
  <tr height="20" class="subttl">
	<td width="75%">Usuário:&nbsp;<span class="selo"><%=Response.Write(Nome)%></td>
	<td width="25%" align="right">Tipo:&nbsp;<span class="selo"><%=Response.Write(Tipo)%></td>
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
</table>
<%rsCadastro.Close%>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
  	<%
	IF Tipo = "INDIVIDUAL" OR Tipo = "GRUPO" THEN
		sSQL = "SELECT * FROM Profissional WHERE profissionalEmail='" & Email02 & "'"
		rsCadastro.Open sSQL, adoCnn, 1, 1
		IF rsCadastro.RecordCount = 0 THEN
		%>
			<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  			<tr>
    		<td class="subttl">&nbsp;ESPECIALIDADES DO PROFISSIONAL</td>
			</tr>
			<tr>
			<td class="aviso">&nbsp;O profissional ainda não informou suas especialidades!</td>
  			</tr>
			</table>
		<%ELSE%>
			<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  			<tr>
    		<td class="subttl">&nbsp;ESPECIALIDADES DO PROFISSIONAL</td>
			</tr>
			<tr>
			<td class="aviso">&nbsp;<%=Response.Write(rsCadastro("profissionalEspecialidades"))%></td>
  			</tr>
			</table>
		<%END IF%>
		<%rsCadastro.Close%>
  		<%SET rsCadastro = Nothing%>
	<%ELSE%>
		<%
		sSQL = "SELECT * FROM Individual WHERE individualEmail='" & Email02 & "'"
		rsCadastro.Open sSQL, adoCnn, 1, 1
		IF rsCadastro.RecordCount = 0 THEN
		%>
			<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  				<tr>
    				<td class="subttl">&nbsp;MAIS SOBRE A PESSOA SELECIONADA</td>
				</tr>
				<tr>
					<td class="aviso">&nbsp;A pessoa ainda não informou mais detalhes!</td>
  				</tr>
			</table>
		<%ELSE%>
			<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  				<tr>
					<td width="100"><p class="rotulo">Altura (M)</p></td>
    				<td width="700" class="aviso">&nbsp;<%=Response.Write(rsCadastro("individualAltura"))%></td>
				</tr>
				<tr>
    				<td width="100"><p class="rotulo">Peso (KG)</p></td>
    				<td width="700" class="aviso">&nbsp;<%=Response.Write(rsCadastro("individualPeso"))%></td>
  				</tr>
			</table>
			<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
			<table width="800" align="center" border="0" cellspacing="0" cellpadding="0">
  			<tr>
    		<td class="subttl">&nbsp;MAIS SOBRE A PESSOA SELECIONADA</td>
			</tr>
			<tr>
			<td class="aviso">&nbsp;<%=Response.Write(rsCadastro("individualSobre"))%></td>
  			</tr>
			</table>
		<%END IF%>
		<%rsCadastro.Close%>
  		<%SET rsCadastro = Nothing%>
	<%END IF%>
  <%adoCnn.Close%>
  <%SET adoCnn = Nothing%>
<p align="center"><img src="../imgs/linha.png" width="800" height="15"></p>
<div class="aviso" align="center">GYM GROUP <%Response.Write("2024" & "-" & Year(Now))%> © Todos os Direitos Reservados</div>
</body>
</html>
